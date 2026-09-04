# Envio Automatizado de Boletos e NFs de Devolução

Script Python que lê uma base de vendedores (Excel), varre uma pasta de PDFs (boletos e notas fiscais de devolução), agrupa os arquivos por cliente e envia um e-mail por cliente com os anexos correspondentes, em cópia para o vendedor e o supervisor responsáveis.

## Requisitos

- Python 3
- Biblioteca `pandas` (e `openpyxl`, usada internamente pelo pandas para ler `.xlsx`)

Instalação:
```
pip install pandas openpyxl
```

## Arquivo de entrada: base Excel

O Excel informado deve conter as seguintes colunas (os nomes originais são renomeados internamente pelo script):

| Coluna original   | Renomeada para     |
|--------------------|---------------------|
| id_vendedor         | id_vendedor         |
| nome                | nome_vendedor        |
| supervisor          | id_supervisor        |
| email_supervisor     | email_supervisor     |
| email_vendedor       | email_vendedor        |

Se alguma dessas colunas estiver faltando, o script interrompe a execução com um erro listando as colunas ausentes.

## Pasta de entrada: PDFs

O script varre a pasta informada e considera apenas arquivos `.pdf`. Cada arquivo é identificado por convenção de nome:

- **Boletos**: o nome do arquivo deve **começar** com o código do cliente (6 dígitos). Ex.: `000000 - Boleto.pdf`
- **Notas fiscais de devolução**: o nome do arquivo deve **começar** com `"NF DE DEV"` e **terminar** com o código do cliente (6 dígitos). Ex.: `NF DE DEV 000000.pdf`

O código do vendedor responsável é deduzido dos **3 primeiros dígitos** do código do cliente (ex.: cliente `000000` → vendedor `000`).

Arquivos cujo nome não se encaixa em nenhum dos dois padrões (código de cliente não identificado) são listados no console como `[SEM CÓDIGO DE CLIENTE]` e ignorados.

## Como o agrupamento e o envio funcionam

1. Todos os PDFs da pasta são agrupados por código de cliente, separando boletos e NFs de devolução.
2. Grupos que têm **NF de devolução mas nenhum boleto** são reportados como `[NF DEV SEM BOLETO]` e não geram envio.
3. Grupos cujo vendedor (pelos 3 primeiros dígitos do código do cliente) não existe na base são reportados como `[VENDEDOR NÃO ENCONTRADO]` e não geram envio.
4. Para os demais grupos (válidos), o script pede o e-mail do cliente **de todos eles primeiro**, um após o outro.
5. Só depois disso ele abre **uma única conexão SMTP** (login único) e envia todos os e-mails nessa mesma conexão — um e-mail por cliente, com os PDFs daquele cliente em anexo.

Cada e-mail enviado tem:
- **Para**: e-mail do cliente informado na hora
- **Cc**: e-mail do vendedor e do supervisor (conforme a base Excel)
- **Assunto**: "Boletos e nota fiscal de devolução"
- **Anexos**: todos os boletos e NFs de devolução daquele cliente

## Executando o script

```
python validacao.py
```

O script vai pedir, nesta ordem:
1. Caminho do arquivo Excel
2. Caminho da pasta com os PDFs
3. E-mail do remetente
4. Senha do remetente (não aparece na tela ao digitar)
5. Em seguida, o e-mail de cada cliente identificado como válido

## Configuração do servidor SMTP

No topo do script:
```python
SMTP_SERVIDOR = "smtps.uhserver.com"
SMTP_PORTA = 465
```
Esses valores são específicos de contas **UOL Host Mail Pro** (domínio próprio hospedado na UOLHOST). Se o e-mail remetente for de outro provedor, esses valores precisam ser ajustados.

## Modo de teste

No topo do script:
```python
MODO_TESTE = False
```
Se `MODO_TESTE = True`, o script executa todo o fluxo (leitura, agrupamento, coleta de e-mails), mas **não abre conexão SMTP nem envia nada** — só imprime no console o que seria enviado para cada cliente.

## Erros conhecidos

- **`SMTPAuthenticationError: (535, ... authentication failed)`**: erro de autenticação no servidor SMTP. Nas contas UOL Host Mail Pro, esse erro pode ocorrer mesmo com login/senha corretos (testados com sucesso no webmail) se a conta não tiver o **envio via aplicativos de terceiros / SMTP externo** liberado no painel administrativo da UOLHOST. Nesse caso, é necessário solicitar essa liberação ao administrador da conta.
