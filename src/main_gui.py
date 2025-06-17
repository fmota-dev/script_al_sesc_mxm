import os
import subprocess
import sys
from pathlib import Path

import customtkinter as ctk
from PIL import Image

from processamento.me import processar_me
from processamento.od import processar_od
from processamento.rf import processar_rf
from utils.helpers import ano_mes_anterior, pedir_codigo_al
from utils.log import escrever_no_log


def resource_path(relative_path):
    """Retorna caminho absoluto para o recurso, funcionando com PyInstaller ou não."""
    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))

    return os.path.join(base_path, relative_path)


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Processador de Arquivos - AL SESC")
        self.iconbitmap(resource_path("src/logo-senac.ico"))
        self.geometry("700x250")
        self.resizable(False, False)
        self.configure(fg_color="#f9f9f9")

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("dark-blue")

        # Variável de pasta
        self.pasta_entrada = ctk.StringVar()

        # Carregando imagens PNG para os botões
        # ATENÇÃO: ajuste os caminhos para suas imagens PNG reais
        self.icon_folder = ctk.CTkImage(
            Image.open(resource_path("src/icons/folder1.png")), size=(20, 20)
        )
        self.icon_run = ctk.CTkImage(
            Image.open(resource_path("src/icons/run1.png")), size=(20, 20)
        )
        self.icon_output = ctk.CTkImage(
            Image.open(resource_path("src/icons/folder1.png")), size=(20, 20)
        )

        self.criar_interface()

    def criar_interface(self):
        # Frame principal com padding
        frame = ctk.CTkFrame(self, fg_color="#f9f9f9", corner_radius=10)
        frame.pack(padx=20, pady=20, fill="both", expand=True)

        # Label instrução
        lbl_instrucao = ctk.CTkLabel(
            frame,
            text="Selecione a pasta com os arquivos ME, OD e RF:",
            anchor="w",
            font=ctk.CTkFont(size=14, weight="normal"),
        )
        lbl_instrucao.pack(fill="x", pady=(0, 10))

        # Frame horizontal para entrada + botão
        frame_entrada = ctk.CTkFrame(
            frame, fg_color="#f9f9f9", corner_radius=0, height=35
        )
        frame_entrada.pack(fill="x")

        # Entry para caminho
        self.entry_pasta = ctk.CTkEntry(
            frame_entrada, textvariable=self.pasta_entrada, width=470, height=35
        )
        self.entry_pasta.pack(side="left", padx=(0, 8), fill="x", expand=True)

        # Botão procurar pasta
        btn_procurar = ctk.CTkButton(
            frame_entrada,
            text="",
            width=35,
            height=35,
            image=self.icon_folder,
            fg_color="#e0e0e0",
            hover_color="#d0d0d0",
            command=self.selecionar_pasta,
        )
        btn_procurar.pack(side="left")

        # Frame botões executar e abrir pasta saída
        frame_botoes = ctk.CTkFrame(frame, fg_color="#f9f9f9", corner_radius=0)
        frame_botoes.pack(pady=20, anchor="w")

        btn_executar = ctk.CTkButton(
            frame_botoes,
            text="Executar",
            width=120,
            height=35,
            image=self.icon_run,
            compound="left",
            command=self.executar,
        )
        btn_executar.pack(side="left", padx=(0, 15))

        btn_abrir_saida = ctk.CTkButton(
            frame_botoes,
            text="Ver arquivos para importação",
            width=160,
            height=35,
            image=self.icon_output,
            compound="left",
            command=self.abrir_pasta_saida,
        )
        btn_abrir_saida.pack(side="left")

        # Label de status
        self.status = ctk.CTkLabel(frame, text="", font=ctk.CTkFont(size=13))
        self.status.pack(pady=(10, 0))

    def selecionar_pasta(self):
        pasta = ctk.filedialog.askdirectory()
        if pasta:
            self.pasta_entrada.set(pasta)

    def abrir_pasta_saida(self):
        pasta_entrada = Path(self.pasta_entrada.get())
        pasta_saida = pasta_entrada / "arquivos_importacao"
        if pasta_saida.exists():
            if os.name == "nt":
                os.startfile(pasta_saida)
            elif sys.platform == "darwin":
                subprocess.run(["open", str(pasta_saida)])
            else:
                subprocess.run(["xdg-open", str(pasta_saida)])
        else:
            self.popup("Aviso", "A pasta de saída ainda não foi gerada.")

    def executar(self):
        pasta_entrada = Path(self.pasta_entrada.get())
        if not pasta_entrada.exists():
            self.popup("Erro", "Pasta de entrada inválida.")
            return

        pasta_saida = pasta_entrada / "arquivos_importacao"
        pasta_logs = pasta_entrada / "logs"
        pasta_saida.mkdir(parents=True, exist_ok=True)
        pasta_logs.mkdir(parents=True, exist_ok=True)

        arquivos = ["ME.xlsx", "OD.xlsx", "RF.xlsx"]
        arquivos_gerados = []

        for arquivo in arquivos:
            entrada = pasta_entrada / arquivo
            if not entrada.exists():
                continue

            saida = pasta_saida / f"{arquivo[:2]}{ano_mes_anterior()}.xlsx"
            log = pasta_logs / f"log_{arquivo[:2].lower()}{ano_mes_anterior()}.txt"
            codigo_al = pedir_codigo_al(arquivo)

            try:
                if arquivo == "ME.xlsx":
                    processar_me(entrada, saida, codigo_al, log)
                elif arquivo == "OD.xlsx":
                    processar_od(entrada, saida, codigo_al, log)
                elif arquivo == "RF.xlsx":
                    processar_rf(entrada, saida, codigo_al, log)

                if saida.exists():
                    arquivos_gerados.append(saida)

                escrever_no_log(f"🆗 Processado: {arquivo}", log)

            except Exception as e:
                escrever_no_log(f"⚠️ Erro: {arquivo} -> {e}", pasta_logs / "erros.txt")

        if arquivos_gerados:
            self.status.configure(
                text=f"✅ {len(arquivos_gerados)} arquivo(s) gerado(s)."
            )
            self.popup_sucesso("Sucesso", "Arquivos gerados com sucesso.")
        else:
            self.status.configure(text="⚠️ Nenhum arquivo gerado.")
            self.popup("Aviso", "Não foram gerados arquivos.")

    def popup(self, titulo, mensagem):
        win = ctk.CTkToplevel(self)
        win.title(titulo)
        win.geometry("350x150")
        win.resizable(False, False)
        win.grab_set()
        win.focus()

        label = ctk.CTkLabel(
            win,
            text=mensagem,
            anchor="center",
            justify="center",
            font=ctk.CTkFont(size=12),
        )
        label.pack(expand=True, padx=20, pady=20)

        btn_fechar = ctk.CTkButton(win, text="Fechar", width=80, command=win.destroy)
        btn_fechar.pack(pady=(0, 20))

        # Centralizar popup
        win.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - (win.winfo_width() // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (win.winfo_height() // 2)
        win.geometry(f"+{x}+{y}")

    def popup_sucesso(self, titulo, mensagem):
        win = ctk.CTkToplevel(self)
        win.title(titulo)
        win.geometry("350x160")
        win.resizable(False, False)
        win.grab_set()
        win.focus()

        label = ctk.CTkLabel(
            win,
            text=mensagem,
            anchor="center",
            justify="center",
            font=ctk.CTkFont(size=12),
        )
        label.pack(expand=True, padx=20, pady=(20, 10))

        frame_botoes = ctk.CTkFrame(win, fg_color="#f9f9f9")
        frame_botoes.pack(pady=(0, 20))

        btn_ver = ctk.CTkButton(
            frame_botoes,
            text="Ver arquivos",
            width=100,
            command=lambda: [self.abrir_pasta_saida(), win.destroy()],
        )
        btn_ver.pack(side="left", padx=(0, 10))

        btn_fechar = ctk.CTkButton(
            frame_botoes, text="Fechar", width=80, command=win.destroy
        )
        btn_fechar.pack(side="left")

        # Centralizar popup
        win.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - (win.winfo_width() // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (win.winfo_height() // 2)
        win.geometry(f"+{x}+{y}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
