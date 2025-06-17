import datetime
import tkinter as tk
from decimal import ROUND_HALF_UP, Decimal

import customtkinter as ctk


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


import customtkinter as ctk


def pedir_codigo_al(arquivo):
    class DialogCodigoAL(ctk.CTkToplevel):
        def __init__(self, parent, arquivo):
            super().__init__(parent)
            self.title("Código AL")
            self.geometry("360x180")
            self.resizable(False, False)
            self.transient(parent)
            self.grab_set()
            self.codigo = None

            self.label = ctk.CTkLabel(
                self, text=f"Digite o código AL para o arquivo {arquivo}:"
            )
            self.label.pack(pady=(20, 12), padx=20)

            self.entry = ctk.CTkEntry(self, width=320)
            self.entry.pack(pady=(0, 12), padx=20)
            self.entry.focus()

            self.lbl_aviso = ctk.CTkLabel(self, text="", text_color="red")
            self.lbl_aviso.pack(pady=(0, 10), padx=20)

            frame_botoes = ctk.CTkFrame(self, fg_color="transparent")
            frame_botoes.pack(pady=5, padx=20, fill="x")

            btn_ok = ctk.CTkButton(
                frame_botoes, text="OK", width=140, command=self.on_ok
            )
            btn_ok.pack(side="left", padx=(0, 15))

            btn_cancel = ctk.CTkButton(
                frame_botoes, text="Cancelar", width=140, command=self.on_cancel
            )
            btn_cancel.pack(side="left")

            self.protocol("WM_DELETE_WINDOW", self.on_cancel)

        def on_ok(self):
            valor = self.entry.get().strip()
            if not valor:
                self.lbl_aviso.configure(text="⚠️ Código AL inválido. Tente novamente.")
                return
            valor_upper = valor.upper()
            if valor_upper.startswith("AL "):
                self.codigo = valor
            elif valor_upper.startswith("AL"):
                self.codigo = "AL " + valor[2:].strip()
            else:
                self.codigo = "AL " + valor
            self.destroy()

        def on_cancel(self):
            self.codigo = None
            self.destroy()

    root = ctk.CTk()
    root.withdraw()

    dialog = DialogCodigoAL(root, arquivo)
    root.wait_window(dialog)

    root.destroy()
    return dialog.codigo
