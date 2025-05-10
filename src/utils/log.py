from datetime import datetime


def escrever_no_log(mensagem, caminho_log):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    mensagem_com_timestamp = f"[{timestamp}] {mensagem}"
    try:
        with open(caminho_log, "a", encoding="utf-8") as log_file:
            log_file.write(mensagem_com_timestamp + "\n")
        print(mensagem_com_timestamp)
    except Exception as e:
        print(f"Erro ao escrever no log: {e}")
