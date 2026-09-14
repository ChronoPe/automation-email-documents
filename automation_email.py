from pathlib import Path
import sys
import re

import getpass
import smtplib
import ssl

from email.message import EmailMessage

import pandas as pd

MODO_TESTE = False

CAMINHO_ASSINATURA = Path(__file__).parent / "assinatura.png"


COLUNAS_ORIGINAIS = {
    "id_vendedor": "id_vendedor",
    "nome": "nome_vendedor",
    "supervisor": "id_supervisor",
    "email_supervisor": "email_supervisor",
    "email_vendedor": "email_vendedor",
}

def ler_base_arquivo(caminho_arquivo: Path) -> pd.DataFrame:
    # Ler o arquivo - EXCEL - e retorna se ele de fato existe.
    caminho = Path(caminho_arquivo)

    if not caminho.is_file():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")

    base = pd.read_excel(caminho, dtype=str)

    base = base.rename(columns=COLUNAS_ORIGINAIS)

    colunas_esperadas = set(COLUNAS_ORIGINAIS.values())
    colunas_ausentes = colunas_esperadas - set(base.columns)

    if colunas_ausentes:
        raise ValueError(f"Colunas ausentes na base: " + ", ".join(sorted(colunas_ausentes)))

    for coluna in base.columns:
        base[coluna] = base[coluna].fillna("").astype(str).str.strip()

    return base

def mostrar_resumo(base: pd.DataFrame) -> None:
    print("Resumo da base de dados:")
    print(f"Registros lidos: {len(base)}")
    print(f"Colunas usadas: {', '.join(base.columns)}")

    sem_email_vendedor = base[base["email_vendedor"] == ""]
    sem_email_supervisor = base[base["email_supervisor"] == ""]

    print(f"Registros sem email de vendedor: {len(sem_email_vendedor)}")
    print(f"Registros sem email de supervisor: {len(sem_email_supervisor)}")

    print("\nPrimeiros 5 registros:")
    print(

        base[["id_vendedor", "nome_vendedor", "email_vendedor", "id_supervisor", "email_supervisor"]].head().to_string(index=False)

    )

    if len(sys.argv) != 2:
        raise SystemExit("Informe o caminho do arquivo Excel como argumento.")

    contatos = ler_base_arquivo(Path(sys.argv[1]))
    mostrar_resumo(contatos)

#--------------------------- 02º parte

def buscar_vendedor_por_id(base: pd.DataFrame, id_vendedor: str):
    id_procurado = str(id_vendedor).strip()

    resultado = base.loc[base["id_vendedor"] == id_procurado]

    if resultado.empty:
        return None

    if len(resultado) > 1:
        raise ValueError(f"Mais de um vendedor encontrado com o ID {id_procurado}.")

    return resultado.iloc[0]

#--------------------------- 03º parte

def extrair_id_vendedor(nome_arquivo: str):
    nome_sem_extensao = Path(nome_arquivo).stem.strip()

    # ^ significa “começo do texto”.
    # (\d{3}) captura exatamente os três primeiros dígitos.
    padrao = r"^(\d{3})"
    resultado = re.search(padrao, nome_sem_extensao)

    if resultado is None:
        return None

    return resultado.group(1)

def listar_pdfs(pasta_entrada: str) -> list[Path]:
    pasta = Path(pasta_entrada)

    if not pasta.is_dir():
        raise NotADirectoryError(f"A pasta especificada não existe: {pasta}")

    pdfs = [arquivo for arquivo in pasta.iterdir()
            if arquivo.is_file() and arquivo.suffix.lower() == ".pdf"]

    return sorted(pdfs)

def extrair_codigo_cliente(nome_arquivo: str):
    nome_sem_extensao = Path(nome_arquivo).stem.strip()

    resultado = re.match(r"^(\d{6})", nome_sem_extensao)

    if resultado is not None:
        return resultado.group(1)

    resultado = re.search(r"(\d{6})\s*(?:\(\d+\))?$", nome_sem_extensao)

    if resultado is not None:
        return resultado.group(1)

    return None

def pedir_email_cliente(codigo_cliente: str) -> str:
    while True:
        email_cliente = input(
            f"Informe o e-mail do cliente {codigo_cliente}: "
        ).strip()

        padrao_email = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"

        if re.fullmatch(padrao_email, email_cliente):
            return email_cliente

        print("E-mail inválido. Tente novamente.")

SMTP_SERVIDOR = "smtps.uhserver.com"
SMTP_PORTA = 465

def montar_mensagem(
    email_remetente: str,
    email_cliente: str,
    vendedor,
    arquivos: list[Path],
) -> EmailMessage:
    mensagem = EmailMessage()

    mensagem["From"] = email_remetente
    mensagem["To"] = email_cliente
    mensagem["Cc"] = (
        f"{vendedor['email_vendedor']}, "
        f"{vendedor['email_supervisor']}"
    )
    mensagem["Subject"] = "Boletos e nota fiscal de devolução"

    mensagem.set_content(
        "Olá,\n\n"
        "Segue em anexo o(s) boleto(s) e documento(s) relacionado(s).\n\n"
        "Atenciosamente."
    )

    mensagem.add_alternative(
        f"""\
<html>
  <body>
    <p>Olá,</p>
    <p>Segue em anexo o(s) boleto(s) e documento(s) relacionado(s).</p>
    <p>Atenciosamente.</p>
    <p><img src="cid:assinatura" alt="Assinatura" style="max-width:350px;"></p>
  </body>
</html>
""",
        subtype="html",
    )

    if CAMINHO_ASSINATURA.is_file():
        with open(CAMINHO_ASSINATURA, "rb") as imagem:
            mensagem.get_payload()[-1].add_related(
                imagem.read(),
                maintype="image",
                subtype="png",
                cid="<assinatura>",
            )
    else:
        print(
            f"[AVISO] Imagem de assinatura não encontrada em: {CAMINHO_ASSINATURA}"
        )

    for arquivo in arquivos:
        with open(arquivo, "rb") as pdf:
            mensagem.add_attachment(
                pdf.read(),
                maintype="application",
                subtype="pdf",
                filename=arquivo.name,
            )

    return mensagem


def enviar_email(smtp: smtplib.SMTP_SSL, mensagem: EmailMessage) -> None:
    smtp.send_message(mensagem)

if __name__ == "__main__":
    caminho_excel = (input("Informe o caminho do arquivo Excel: ")).strip().strip('"').strip("'")
    pasta_entrada = (input("Cole o caminho da pasta com os PDFs: ")).strip().strip('"').strip("'")

    contatos = ler_base_arquivo(Path(caminho_excel))
    pdfs = listar_pdfs(pasta_entrada)

    email_remetente = input("Informe o e-mail do remetente: ").strip()
    senha_remetente = getpass.getpass("Informe a senha do e-mail do remetente: ").strip()

    print(f'Autenticação realizada com sucesso para o e-mail: {email_remetente}\n')

    grupos_por_cliente = {}

    for pdf in pdfs:
        nome = pdf.name

        nome_maiusculo = nome.upper()

        if nome_maiusculo.startswith("NF DE DEV") or nome_maiusculo.startswith("NF CL"):
            tipo_arquivo = "nf_dev"
        else:
            tipo_arquivo = "boleto"

        codigo_cliente = extrair_codigo_cliente(nome)

        if codigo_cliente is None:
            print(f"[SEM CÓDIGO DE CLIENTE] {nome}")
            continue

        if codigo_cliente not in grupos_por_cliente:
            grupos_por_cliente[codigo_cliente] = {
                "boletos": [],
                "nfs_dev": [],
            }

        if tipo_arquivo == "boleto":
            grupos_por_cliente[codigo_cliente]["boletos"].append(pdf)
        else:
            grupos_por_cliente[codigo_cliente]["nfs_dev"].append(pdf)

    # 1ª etapa: filtra os grupos válidos (com boleto e vendedor encontrado)
    grupos_validos = {}

    for codigo_cliente, grupo in sorted(grupos_por_cliente.items()):
        id_vendedor = codigo_cliente[:3]

        vendedor = buscar_vendedor_por_id(contatos, id_vendedor)

        if not grupo["boletos"]:
            print(f"[SEM BOLETO - envio será feito só com NF] Cliente: {codigo_cliente}")
            for nf_dev in grupo["nfs_dev"]:
                print(f"  {nf_dev.name}")

        if vendedor is None:
            print(
                f"[VENDEDOR NÃO ENCONTRADO] "
                f"Cliente: {codigo_cliente} | ID: {id_vendedor}"
            )
            continue

        grupos_validos[codigo_cliente] = {
            "grupo": grupo,
            "vendedor": vendedor,
        }

    # 2ª etapa: pede o e-mail de todos os clientes válidos, de uma vez
    emails_por_cliente = {}

    for codigo_cliente in grupos_validos:
        emails_por_cliente[codigo_cliente] = pedir_email_cliente(codigo_cliente)

    # 3ª etapa: envia tudo numa única conexão SMTP
    if MODO_TESTE:
        for codigo_cliente in grupos_validos:
            print(f"\n[PRONTO PARA ENVIO] Cliente: {codigo_cliente}")
            print(f"E-mail do cliente: {emails_por_cliente[codigo_cliente]}")
            print("MODO DE TESTE ATIVADO: E-mail não enviado.\n")
    else:
        contexto_ssl = ssl.create_default_context()

        with smtplib.SMTP_SSL(
            SMTP_SERVIDOR,
            SMTP_PORTA,
            context=contexto_ssl,
        ) as smtp:
            smtp.login(email_remetente, senha_remetente)

            for codigo_cliente, dados in grupos_validos.items():
                grupo = dados["grupo"]
                vendedor = dados["vendedor"]
                email_cliente = emails_por_cliente[codigo_cliente]

                print(f"\n[PRONTO PARA ENVIO] Cliente: {codigo_cliente}")
                print(f"E-mail do cliente: {email_cliente}")

                arquivos_para_enviar = grupo["boletos"] + grupo["nfs_dev"]

                mensagem = montar_mensagem(
                    email_remetente,
                    email_cliente,
                    vendedor,
                    arquivos_para_enviar,
                )

                try:
                    enviar_email(smtp, mensagem)
                except (smtplib.SMTPException, OSError) as erro:
                    print(f"[ERRO AO ENVIAR] Cliente {codigo_cliente}: {erro}\n")
                    continue

                print("[E-MAIL ENVIADO COM SUCESSO]\n")
            