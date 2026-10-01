#py -3.14 auto.py
import json
import os
import threading
import time
import ctypes
import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
from playwright.sync_api import sync_playwright

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class AppRPA(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.title("Automação de Cadastro")

        largura_janela = 550
        altura_janela = 550

        self.centralizar_janela(largura_janela, altura_janela)

        self.resizable(False, False)
        self.configure(fg_color="#121318")

        self.caminho_arquivo = ""

        self.card_cred = ctk.CTkFrame(
            self,
            corner_radius=10,
            fg_color="#1A1B23",
            border_color="#2D303E",
            border_width=1,
        )
        self.card_cred.pack(pady=(25, 10), padx=25, fill="x")

        self.lbl_cred = ctk.CTkLabel(
            self.card_cred,
            text="Credenciais de Acesso",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#E2E8F0",
        )
        self.lbl_cred.pack(pady=(12, 8), padx=15, anchor="w")

        self.input_url = ctk.CTkEntry(
            self.card_cred,
            placeholder_text="URL do Sistema",
            height=38,
            border_width=1,
            border_color="#374151",
            fg_color="#111827",
            text_color="#F3F4F6",
            corner_radius=6,
        )
        self.input_url.pack(pady=5, padx=15, fill="x")

        self.input_email = ctk.CTkEntry(
            self.card_cred,
            placeholder_text="E-mail de Login",
            height=38,
            border_width=1,
            border_color="#374151",
            fg_color="#111827",
            text_color="#F3F4F6",
            corner_radius=6,
        )
        self.input_email.pack(pady=5, padx=15, fill="x")

        self.input_senha = ctk.CTkEntry(
            self.card_cred,
            placeholder_text="Senha",
            show="*",
            height=38,
            border_width=1,
            border_color="#374151",
            fg_color="#111827",
            text_color="#F3F4F6",
            corner_radius=6,
        )
        self.input_senha.pack(pady=(5, 12), padx=15, fill="x")

        self.card_file = ctk.CTkFrame(
            self,
            corner_radius=10,
            fg_color="#1A1B23",
            border_color="#2D303E",
            border_width=1,
        )
        self.card_file.pack(pady=10, padx=25, fill="x")

        self.lbl_file = ctk.CTkLabel(
            self.card_file,
            text="Base de Dados (Excel ou CSV)",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#E2E8F0",
        )
        self.lbl_file.pack(pady=(12, 8), padx=15, anchor="w")

        self.file_subframe = ctk.CTkFrame(
            self.card_file, fg_color="transparent"
        )
        self.file_subframe.pack(pady=(0, 12), padx=15, fill="x")

        self.btn_selecionar = ctk.CTkButton(
            self.file_subframe,
            text="Selecionar Planilha",
            command=self.selecionar_arquivo,
            height=35,
            width=140,
            corner_radius=6,
            fg_color="#1E3A8A",
            hover_color="#1E40AF",
            text_color="#FFFFFF",
        )
        self.btn_selecionar.pack(side="left")

        self.lbl_arquivo = ctk.CTkLabel(
            self.file_subframe,
            text="Nenhum arquivo selecionado",
            text_color="#9CA3AF",
            font=ctk.CTkFont(size=12),
        )
        self.lbl_arquivo.pack(side="left", padx=15)

        self.opt_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.opt_frame.pack(pady=5, padx=25, fill="x")

        self.chk_oculto = ctk.CTkCheckBox(
            self.opt_frame,
            text="Executar em segundo plano (Headless)",
            font=ctk.CTkFont(size=12),
            text_color="#D1D5DB",
            fg_color="#1D4ED8",
            hover_color="#1E40AF",
        )
        self.chk_oculto.pack(anchor="w", pady=(0, 10))

        self.progress_bar = ctk.CTkProgressBar(
            self.opt_frame,
            height=10,
            corner_radius=5,
            fg_color="#1F2937",
            progress_color="#2563EB",
        )
        self.progress_bar.pack(fill="x")
        self.progress_bar.set(0)

        self.btn_iniciar = ctk.CTkButton(
            self,
            text="INICIAR AUTOMATIZAÇÃO",
            command=self.iniciar_thread_rpa,
            height=46,
            font=ctk.CTkFont(size=14, weight="bold"),
            corner_radius=8,
            fg_color="#1D4ED8",
            hover_color="#2563EB",
            text_color="#FFFFFF",
        )
        self.btn_iniciar.pack(pady=15, padx=25, fill="x")

        self.log_textbox = ctk.CTkTextbox(
            self,
            height=120,
            corner_radius=8,
            fg_color="#1A1B23",
            border_color="#2D303E",
            border_width=1,
            text_color="#D1D5DB",
            font=ctk.CTkFont(family="Consolas", size=11),
        )
        self.log_textbox.pack(pady=(0, 20), padx=25, fill="x")
        self.log("Aguardando início das operações...")

    def centralizar_janela(self, largura, altura):

        class RECT(ctypes.Structure):
            _fields_ = [
                ("left", ctypes.c_long),
                ("top", ctypes.c_long),
                ("right", ctypes.c_long),
                ("bottom", ctypes.c_long),
            ]

        SPI_GETWORKAREA = 48

        rect = RECT()
        ctypes.windll.user32.SystemParametersInfoW(
            SPI_GETWORKAREA,
            0,
            ctypes.byref(rect),
            0,
        )

        largura_util = rect.right - rect.left
        altura_util = rect.bottom - rect.top

        pos_x = rect.left + (largura_util - largura) // 2
        pos_y = rect.top + (altura_util - altura) // 2

        self.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

    def log(self, mensagem):
        self.log_textbox.insert("end", f"> {mensagem}\n")
        self.log_textbox.see("end")

    def carregar_config_inicial(self):
        try:
            with open("config.json", "r", encoding="utf-8") as f:
                config = json.load(f)
                self.input_url.insert(0, config["configuracao"]["url"])
                self.input_email.insert(0, config["credenciais"]["email"])
                self.input_senha.insert(0, config["credenciais"]["senha"])
        except Exception:
            pass

    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
            filetypes=[
                ("Planilhas", "*.csv;*.xlsx;*.xls"),
                ("Arquivos CSV", "*.csv"),
                ("Arquivos Excel", "*.xlsx;*.xls"),
            ]
        )

        if caminho:
            self.caminho_arquivo = caminho
            nome_arquivo = caminho.split("/")[-1]

            self.lbl_arquivo.configure(
                text=f"✔ {nome_arquivo}",
                text_color="#60A5FA"
            )

            self.log(f"Arquivo carregado: {nome_arquivo}")

    def iniciar_thread_rpa(self):

        if not self.caminho_arquivo:
            messagebox.showwarning(
                "Atenção",
                "Por favor, selecione uma planilha primeiro!"
            )
            return

        if (
            not self.input_url.get()
            or not self.input_email.get()
            or not self.input_senha.get()
        ):
            messagebox.showwarning(
                "Atenção",
                "Preencha a URL, o e-mail e a senha antes de iniciar."
            )
            return

        self.btn_iniciar.configure(
            state="disabled",
            fg_color="#374151"
        )

        self.progress_bar.set(0)

        threading.Thread(
            target=self.executar_rpa,
            daemon=True
        ).start()

    def executar_rpa(self):

        url = self.input_url.get()
        email = self.input_email.get()
        senha = self.input_senha.get()

        modo_oculto = self.chk_oculto.get() == 1

        try:

            self.log("Lendo arquivo de dados...")

            extensao = os.path.splitext(
                self.caminho_arquivo
            )[1].lower()

            if extensao in [".xlsx", ".xls"]:
                tabela = pd.read_excel(
                    self.caminho_arquivo
                )
            else:
                tabela = pd.read_csv(
                    self.caminho_arquivo
                )

            total = len(tabela)

            with sync_playwright() as p:

                self.log("Iniciando navegador...")

                browser = p.chromium.launch(
                    headless=modo_oculto,
                    slow_mo=100 if not modo_oculto else 0
                )

                page = browser.new_page()

                self.log("Acessando sistema...")

                if not url.startswith(
                    ("http://", "https://")
                ):
                    url = "https://" + url

                page.goto(url)

                page.locator("#email").fill(email)
                page.locator("#password").fill(senha)

                page.locator("#password").press("Enter")

                time.sleep(2)

                self.log(
                    f"Processando {total} produtos..."
                )

                for index, produto in tabela.iterrows():

                    progresso = (index + 1) / total

                    self.progress_bar.set(
                        progresso
                    )

                    self.log(
                        f"[{index + 1}/{total}] "
                        f"Cadastrando: {produto['codigo']}"
                    )

                    page.locator("#codigo").fill(
                        str(produto["codigo"])
                    )

                    page.locator("#marca").fill(
                        str(produto["marca"])
                    )

                    page.locator("#tipo").fill(
                        str(produto["tipo"])
                    )

                    page.locator("#categoria").fill(
                        str(produto["categoria"])
                    )

                    page.locator("#preco_unitario").fill(
                        str(produto["preco_unitario"])
                    )

                    page.locator("#custo").fill(
                        str(produto["custo"])
                    )

                    obs = str(produto["obs"])

                    page.locator("#obs").fill(
                        obs if obs != "nan" else ""
                    )

                    page.locator("#obs").press("Enter")

                    if page.locator(
                        "button:has-text('Enviar')"
                    ).is_visible():

                        page.locator(
                            "button:has-text('Enviar')"
                        ).click(force=True)

                browser.close()

                self.log(
                    "--- PROCESSO CONCLUÍDO COM SUCESSO! ---"
                )

                messagebox.showinfo(
                    "Sucesso",
                    "Todos os produtos foram cadastrados!"
                )

        except Exception as e:

            self.log(
                f"ERRO: {str(e)}"
            )

            messagebox.showerror(
                "Erro",
                f"Ocorreu um erro durante o processo:\n{e}"
            )

        finally:

            self.btn_iniciar.configure(
                state="normal",
                fg_color="#1D4ED8"
            )


if __name__ == "__main__":
    app = AppRPA()
    app.mainloop()