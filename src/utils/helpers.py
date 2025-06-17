import datetime
import math
from decimal import ROUND_HALF_UP, Decimal


def truncar_se_mais_de_duas_casas(valor, acumulador):
    partes = str(valor).split(".")
    if len(partes) == 2 and len(partes[1]) > 2:
        truncado = int(valor * 100) / 100
        acumulador["truncado"] += valor - truncado
        return truncado
    return valor


def arredondar(valor):
    return float(Decimal(valor).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def ultimo_dia_mes_anterior():
    hoje = datetime.datetime.now()
    primeiro_dia_mes = datetime.datetime(hoje.year, hoje.month, 1)
    ultimo_dia = primeiro_dia_mes - datetime.timedelta(days=1)
    return ultimo_dia.strftime("%Y%m%d")


def ano_mes_anterior():
    hoje = datetime.datetime.now()
    mes = hoje.month - 1 or 12
    ano = hoje.year if hoje.month > 1 else hoje.year - 1
    return f"{ano}{mes:02d}"


def nome_documento(tipo: str) -> str:
    hoje = datetime.datetime.now()
    mes = hoje.month - 1 or 12
    ano = hoje.year if hoje.month > 1 else hoje.year - 1
    return f"{tipo.upper()}{mes:02d}{ano % 100:02d}"


def formatar_historico(codigo_al, area):
    ano_mes = ano_mes_anterior()
    return f"{area} ({codigo_al} SESC) {ano_mes[4:]}/{ano_mes[:4]}"


def pedir_codigo_al(arquivo):
    while True:
        codigo_al = input(f"Digite o código AL para o arquivo {arquivo}: ").strip()
        if not codigo_al:
            print("⚠️ Código AL inválido. Tente novamente.")
            continue
        if codigo_al.upper().startswith("AL "):
            return codigo_al
        elif codigo_al.upper().startswith("AL"):
            return "AL " + codigo_al[2:].strip()
        else:
            return "AL " + codigo_al
