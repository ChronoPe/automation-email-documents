"""
Front-end (CustomTkinter) para o sistema de envio de boletos e NFs de devolução.

Este arquivo NÃO altera automation_email.py — apenas importa e reutiliza as
funções e constantes já existentes lá (leitura da base, listagem de PDFs,
busca de vendedor, montagem e envio de e-mail). A lógica de agrupamento por
cliente, que no back-end vivia solta dentro do bloco "if __name__ == '__main__':",
foi replicada aqui em `agrupar_pdfs_por_cliente`, pois lá ela não existe como
função reutilizável.
"""

import queue
import re
import smtplib
import ssl
import threading
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

import automation_email as backend

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

EMAIL_REGEX = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"


def agrupar_pdfs_por_cliente(base, pasta_pdfs):
    """Replica o agrupamento feito no __main__ de automation_email.py."""
    pdfs = backend.listar_pdfs(pasta_pdfs)

    grupos_por_cliente = {}
    sem_codigo = []

    for pdf in pdfs:
        nome = pdf.name
        nome_maiusculo = nome.upper()

        if nome_maiusculo.startswith("NF DE DEV") or nome_maiusculo.startswith("NF CL"):
            tipo_arquivo = "nf_dev"
        else:
            tipo_arquivo = "boleto"

        codigo_cliente = backend.extrair_codigo_cliente(nome)

        if codigo_cliente is None:
            sem_codigo.append(nome)
            continue

        grupos_por_cliente.setdefault(codigo_cliente, {"boletos": [], "nfs_dev": []})

        if tipo_arquivo == "boleto":
            grupos_por_cliente[codigo_cliente]["boletos"].append(pdf)
        else:
            grupos_por_cliente[codigo_cliente]["nfs_dev"].append(pdf)

    grupos_validos = {}
    sem_boleto = []
    vendedor_nao_encontrado = []

    for codigo_cliente, grupo in sorted(grupos_por_cliente.items()):
        id_vendedor = codigo_cliente[:3]
        vendedor = backend.buscar_vendedor_por_id(base, id_vendedor)

        if not grupo["boletos"]:
            sem_boleto.append(codigo_cliente)

        if vendedor is None:
            vendedor_nao_encontrado.append((codigo_cliente, id_vendedor))
            continue

        grupos_validos[codigo_cliente] = {"grupo": grupo, "vendedor": vendedor}

    return {
        "total_pdfs": len(pdfs),
        "grupos_validos": grupos_validos,
        "sem_codigo": sem_codigo,
        "sem_boleto": sem_boleto,
        "vendedor_nao_encontrado": vendedor_nao_encontrado,
    }


class BoletoApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Envio de Boletos e NFs de Devolução")
        self.geometry("1000x700")
        self.minsize(880, 600)

        self.base = None
        self.resultado = None
        self.email_entries = {}
        self.log_queue = queue.Queue()

        self.tabview = ctk.CTkTabview(self, width=960, height=660)
        self.tabview.pack(padx=20, pady=20, fill="both", expand=True)

        self.tab_config = self.tabview.add("1. Configuração")
        self.tab_revisao = self.tabview.add("2. Revisão")
        self.tab_envio = self.tabview.add("3. Envio")

        self._build_tab_config()
        self._build_tab_revisao_vazia()
        self._build_tab_envio_vazia()

    # ------------------------------------------------------------------ #
    # Aba 1 - Configuração
    # ------------------------------------------------------------------ #
    def _build_tab_config(self):
        frame = ctk.CTkFrame(self.tab_config, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(frame, text="Base de vendedores (Excel)", anchor="w").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 4)
        )
        linha_excel = ctk.CTkFrame(frame, fg_color="transparent")
        linha_excel.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 16))
        linha_excel.grid_columnconfigure(0, weight=1)

        self.entry_excel = ctk.CTkEntry(linha_excel, placeholder_text="Caminho do arquivo .xlsx")
        self.entry_excel.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ctk.CTkButton(linha_excel, text="Procurar...", width=110, command=self._selecionar_excel).grid(
            row=0, column=1
        )

        ctk.CTkLabel(frame, text="Pasta com os PDFs (boletos e NFs de devolução)", anchor="w").grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(0, 4)
        )
        linha_pasta = ctk.CTkFrame(frame, fg_color="transparent")
        linha_pasta.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 16))
        linha_pasta.grid_columnconfigure(0, weight=1)

        self.entry_pasta = ctk.CTkEntry(linha_pasta, placeholder_text="Caminho da pasta com os PDFs")
        self.entry_pasta.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ctk.CTkButton(linha_pasta, text="Procurar...", width=110, command=self._selecionar_pasta).grid(
            row=0, column=1
        )

        ctk.CTkLabel(frame, text="E-mail do remetente", anchor="w").grid(
            row=4, column=0, columnspan=2, sticky="w", pady=(0, 4)
        )
        self.entry_remetente = ctk.CTkEntry(frame, placeholder_text="seu.email@dominio.com.br")
        self.entry_remetente.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(0, 16))

        ctk.CTkLabel(frame, text="Senha do e-mail do remetente", anchor="w").grid(
            row=6, column=0, columnspan=2, sticky="w", pady=(0, 4)
        )
        self.entry_senha = ctk.CTkEntry(frame, placeholder_text="Senha", show="*")
        self.entry_senha.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(0, 16))

        self.var_modo_teste = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            frame,
            text="Modo teste (monta tudo, mas não conecta ao SMTP nem envia nada)",
            variable=self.var_modo_teste,
        ).grid(row=8, column=0, columnspan=2, sticky="w", pady=(0, 20))

        self.btn_processar = ctk.CTkButton(
            frame, text="Processar base e PDFs", command=self._iniciar_processamento
        )
        self.btn_processar.grid(row=9, column=0, columnspan=2, sticky="w")

        self.label_status_config = ctk.CTkLabel(frame, text="", text_color="tomato", anchor="w", justify="left")
        self.label_status_config.grid(row=10, column=0, columnspan=2, sticky="w", pady=(12, 0))

    def _selecionar_excel(self):
        caminho = filedialog.askopenfilename(
            title="Selecione a base de vendedores",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")],
        )
        if caminho:
            self.entry_excel.delete(0, "end")
            self.entry_excel.insert(0, caminho)

    def _selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta com os PDFs")
        if pasta:
            self.entry_pasta.delete(0, "end")
            self.entry_pasta.insert(0, pasta)

    def _iniciar_processamento(self):
        caminho_excel = self.entry_excel.get().strip().strip('"').strip("'")
        pasta_pdfs = self.entry_pasta.get().strip().strip('"').strip("'")

        if not caminho_excel or not pasta_pdfs:
            self.label_status_config.configure(text="Informe o arquivo Excel e a pasta de PDFs.")
            return

        self.label_status_config.configure(text="")
        self.btn_processar.configure(state="disabled", text="Processando...")

        threading.Thread(
            target=self._thread_processar, args=(caminho_excel, pasta_pdfs), daemon=True
        ).start()

    def _thread_processar(self, caminho_excel, pasta_pdfs):
        try:
            base = backend.ler_base_arquivo(Path(caminho_excel))
            resultado = agrupar_pdfs_por_cliente(base, pasta_pdfs)
        except Exception as erro:
            self.after(0, self._processamento_erro, str(erro))
            return

        self.after(0, self._processamento_concluido, base, resultado)

    def _processamento_erro(self, mensagem):
        self.btn_processar.configure(state="normal", text="Processar base e PDFs")
        self.label_status_config.configure(text=f"Erro: {mensagem}")

    def _processamento_concluido(self, base, resultado):
        self.btn_processar.configure(state="normal", text="Processar base e PDFs")
        self.base = base
        self.resultado = resultado

        if not resultado["grupos_validos"]:
            self.label_status_config.configure(
                text="Processado, mas nenhum cliente válido foi encontrado (veja a aba Revisão)."
            )
        else:
            self.label_status_config.configure(text="", text_color="tomato")

        self._build_tab_revisao_conteudo()
        self.tabview.set("2. Revisão")

    # ------------------------------------------------------------------ #
    # Aba 2 - Revisão
    # ------------------------------------------------------------------ #
    def _build_tab_revisao_vazia(self):
        self.revisao_placeholder = ctk.CTkLabel(
            self.tab_revisao,
            text="Processe a base e os PDFs na aba anterior para revisar os clientes aqui.",
            text_color="gray",
        )
        self.revisao_placeholder.pack(padx=10, pady=20)

    def _build_tab_revisao_conteudo(self):
        for widget in self.tab_revisao.winfo_children():
            widget.destroy()

        self.email_entries = {}
        resultado = self.resultado

        frame = ctk.CTkFrame(self.tab_revisao, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        frame.grid_rowconfigure(2, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        resumo = (
            f"PDFs encontrados: {resultado['total_pdfs']}    |    "
            f"Clientes válidos para envio: {len(resultado['grupos_validos'])}    |    "
            f"Sem código de cliente: {len(resultado['sem_codigo'])}    |    "
            f"Vendedor não encontrado: {len(resultado['vendedor_nao_encontrado'])}"
        )
        ctk.CTkLabel(frame, text=resumo, anchor="w", justify="left").grid(
            row=0, column=0, sticky="w", pady=(0, 10)
        )

        avisos = []
        for nome in resultado["sem_codigo"]:
            avisos.append(f"[SEM CÓDIGO DE CLIENTE] {nome}")
        for codigo, id_vendedor in resultado["vendedor_nao_encontrado"]:
            avisos.append(f"[VENDEDOR NÃO ENCONTRADO] Cliente: {codigo} | ID: {id_vendedor}")

        if avisos:
            caixa_avisos = ctk.CTkTextbox(frame, height=90)
            caixa_avisos.grid(row=1, column=0, sticky="ew", pady=(0, 10))
            caixa_avisos.insert("end", "\n".join(avisos))
            caixa_avisos.configure(state="disabled")

        lista = ctk.CTkScrollableFrame(frame, label_text="Clientes válidos — informe o e-mail de cada um")
        lista.grid(row=2, column=0, sticky="nsew")
        lista.grid_columnconfigure(3, weight=1)

        cabecalho = ["Cliente", "Vendedor", "Arquivos", "E-mail do cliente"]
        for col, texto in enumerate(cabecalho):
            ctk.CTkLabel(lista, text=texto, font=ctk.CTkFont(weight="bold")).grid(
                row=0, column=col, sticky="w", padx=6, pady=(0, 6)
            )

        sem_boleto_set = set(resultado["sem_boleto"])

        for i, (codigo_cliente, dados) in enumerate(resultado["grupos_validos"].items(), start=1):
            grupo = dados["grupo"]
            vendedor = dados["vendedor"]
            n_arquivos = len(grupo["boletos"]) + len(grupo["nfs_dev"])

            rotulo_cliente = codigo_cliente
            if codigo_cliente in sem_boleto_set:
                rotulo_cliente += "  ⚠ só NF, sem boleto"

            ctk.CTkLabel(lista, text=rotulo_cliente, anchor="w").grid(
                row=i, column=0, sticky="w", padx=6, pady=3
            )
            ctk.CTkLabel(lista, text=vendedor.get("nome_vendedor", ""), anchor="w").grid(
                row=i, column=1, sticky="w", padx=6, pady=3
            )
            ctk.CTkLabel(lista, text=str(n_arquivos), anchor="w").grid(
                row=i, column=2, sticky="w", padx=6, pady=3
            )

            entry = ctk.CTkEntry(lista, placeholder_text="cliente@dominio.com.br")
            entry.grid(row=i, column=3, sticky="ew", padx=6, pady=3)
            self.email_entries[codigo_cliente] = entry

        botoes = ctk.CTkFrame(frame, fg_color="transparent")
        botoes.grid(row=3, column=0, sticky="w", pady=(12, 0))

        self.label_status_revisao = ctk.CTkLabel(botoes, text="", text_color="tomato")
        self.label_status_revisao.pack(side="left", padx=(0, 12))

        estado_botao = "normal" if resultado["grupos_validos"] else "disabled"
        ctk.CTkButton(
            botoes, text="Validar e ir para envio", state=estado_botao, command=self._validar_emails_e_avancar
        ).pack(side="left")

    def _validar_emails_e_avancar(self):
        faltando = []
        invalidos = []

        for codigo_cliente, entry in self.email_entries.items():
            email = entry.get().strip()
            if not email:
                faltando.append(codigo_cliente)
            elif not re.fullmatch(EMAIL_REGEX, email):
                invalidos.append(codigo_cliente)

        if faltando or invalidos:
            partes = []
            if faltando:
                partes.append(f"sem e-mail: {', '.join(faltando)}")
            if invalidos:
                partes.append(f"e-mail inválido: {', '.join(invalidos)}")
            self.label_status_revisao.configure(text="Corrija antes de continuar — " + "; ".join(partes))
            return

        self.label_status_revisao.configure(text="")
        self._build_tab_envio_conteudo()
        self.tabview.set("3. Envio")

    # ------------------------------------------------------------------ #
    # Aba 3 - Envio
    # ------------------------------------------------------------------ #
    def _build_tab_envio_vazia(self):
        self.envio_placeholder = ctk.CTkLabel(
            self.tab_envio,
            text="Revise os clientes e valide os e-mails na aba anterior para liberar o envio.",
            text_color="gray",
        )
        self.envio_placeholder.pack(padx=10, pady=20)

    def _build_tab_envio_conteudo(self):
        for widget in self.tab_envio.winfo_children():
            widget.destroy()

        frame = ctk.CTkFrame(self.tab_envio, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        frame.grid_rowconfigure(2, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        modo_teste = self.var_modo_teste.get()
        total = len(self.resultado["grupos_validos"])
        texto_modo = "MODO TESTE — nenhum e-mail será enviado de verdade" if modo_teste else "ENVIO REAL"
        ctk.CTkLabel(
            frame,
            text=f"{total} e-mail(s) prontos para envio.    |    {texto_modo}",
            anchor="w",
        ).grid(row=0, column=0, sticky="w", pady=(0, 10))

        self.btn_enviar = ctk.CTkButton(frame, text="Enviar e-mails agora", command=self._confirmar_envio)
        self.btn_enviar.grid(row=1, column=0, sticky="w", pady=(0, 10))

        self.progress_envio = ctk.CTkProgressBar(frame)
        self.progress_envio.set(0)
        self.progress_envio.grid(row=1, column=0, sticky="e")

        self.log_envio = ctk.CTkTextbox(frame)
        self.log_envio.grid(row=2, column=0, sticky="nsew")
        self.log_envio.configure(state="disabled")

    def _confirmar_envio(self):
        modo_teste = self.var_modo_teste.get()

        if not modo_teste:
            confirmar = messagebox.askyesno(
                "Confirmar envio",
                f"Isso vai enviar {len(self.resultado['grupos_validos'])} e-mail(s) reais agora. Continuar?",
            )
            if not confirmar:
                return

        email_remetente = self.entry_remetente.get().strip()
        senha = self.entry_senha.get()

        if not modo_teste and (not email_remetente or not senha):
            messagebox.showerror("Erro", "Informe o e-mail e a senha do remetente na aba Configuração.")
            return

        self.btn_enviar.configure(state="disabled", text="Enviando...")
        self.progress_envio.set(0)
        self._log_envio_limpar()

        threading.Thread(
            target=self._thread_enviar, args=(email_remetente, senha, modo_teste), daemon=True
        ).start()
        self.after(100, self._poll_log_queue)

    def _log_envio_limpar(self):
        self.log_envio.configure(state="normal")
        self.log_envio.delete("1.0", "end")
        self.log_envio.configure(state="disabled")

    def _log_envio_escrever(self, linha):
        self.log_envio.configure(state="normal")
        self.log_envio.insert("end", linha + "\n")
        self.log_envio.see("end")
        self.log_envio.configure(state="disabled")

    def _thread_enviar(self, email_remetente, senha, modo_teste):
        grupos_validos = self.resultado["grupos_validos"]
        total = len(grupos_validos)
        enviados = 0
        falhas = 0

        if modo_teste:
            for codigo_cliente in grupos_validos:
                email_cliente = self.email_entries[codigo_cliente].get().strip()
                self.log_queue.put(
                    ("log", f"[TESTE] Cliente {codigo_cliente} -> {email_cliente} (não enviado)")
                )
                enviados += 1
                self.log_queue.put(("progress", enviados / total))
            self.log_queue.put(("done", enviados, falhas))
            return

        try:
            contexto_ssl = ssl.create_default_context()
            with smtplib.SMTP_SSL(backend.SMTP_SERVIDOR, backend.SMTP_PORTA, context=contexto_ssl) as smtp:
                smtp.login(email_remetente, senha)
                self.log_queue.put(("log", f"Autenticação realizada com sucesso para: {email_remetente}"))

                for codigo_cliente, dados in grupos_validos.items():
                    grupo = dados["grupo"]
                    vendedor = dados["vendedor"]
                    email_cliente = self.email_entries[codigo_cliente].get().strip()
                    arquivos = grupo["boletos"] + grupo["nfs_dev"]

                    mensagem = backend.montar_mensagem(email_remetente, email_cliente, vendedor, arquivos)

                    try:
                        backend.enviar_email(smtp, mensagem)
                        enviados += 1
                        self.log_queue.put(("log", f"[OK] Cliente {codigo_cliente} -> {email_cliente}"))
                    except (smtplib.SMTPException, OSError) as erro:
                        falhas += 1
                        self.log_queue.put(("log", f"[ERRO] Cliente {codigo_cliente}: {erro}"))

                    self.log_queue.put(("progress", (enviados + falhas) / total))
        except Exception as erro:
            self.log_queue.put(("auth_erro", str(erro)))
            return

        self.log_queue.put(("done", enviados, falhas))

    def _poll_log_queue(self):
        terminou = False

        while True:
            try:
                item = self.log_queue.get_nowait()
            except queue.Empty:
                break

            tipo = item[0]
            if tipo == "log":
                self._log_envio_escrever(item[1])
            elif tipo == "progress":
                self.progress_envio.set(item[1])
            elif tipo == "auth_erro":
                self._log_envio_escrever(f"[ERRO DE CONEXÃO/AUTENTICAÇÃO] {item[1]}")
                messagebox.showerror("Erro ao enviar", item[1])
                terminou = True
            elif tipo == "done":
                enviados, falhas = item[1], item[2]
                self._log_envio_escrever(f"\nConcluído: {enviados} enviado(s), {falhas} falha(s).")
                messagebox.showinfo("Envio concluído", f"{enviados} enviado(s), {falhas} falha(s).")
                terminou = True

        if terminou:
            self.btn_enviar.configure(state="normal", text="Enviar e-mails agora")
        else:
            self.after(100, self._poll_log_queue)


if __name__ == "__main__":
    app = BoletoApp()
    app.mainloop()
