import requests
import pandas as pd
import re
import time
import os
from datetime import datetime
from msal import ConfidentialClientApplication

# ==================================================
# CONFIGURAÇÕES GERAIS (GITHUB-SAFE)
# ==================================================

# >>> CONFIGURE VIA VARIÁVEIS DE AMBIENTE <<<
CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
TENANT_ID = os.getenv("AZURE_TENANT_ID")
MAILBOX = os.getenv("MAILBOX_EMAIL")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

# Filtro de datas
DATA_INICIO = "2025-01-01"
DATA_FIM = "2025-08-01"

palavra_chave = input("Digite a palavra para filtrar os e-mails: ").lower().strip()

# ==================================================
# DIRETÓRIOS
# ==================================================

PASTA_BASE = r"C:\chamados"
PASTA_ENTRADA = os.path.join(PASTA_BASE, "chamadosFiltrados")
PASTA_SAIDA = os.path.join(PASTA_BASE, "chamadosOrganizados")
LOG_FOLDER = os.path.join(PASTA_BASE, "logs")

os.makedirs(PASTA_ENTRADA, exist_ok=True)
os.makedirs(PASTA_SAIDA, exist_ok=True)
os.makedirs(LOG_FOLDER, exist_ok=True)

palavra_arquivo = re.sub(r"[^\w\-]", "_", palavra_chave)
NOME_EXCEL = f"chamados_filtrados_{palavra_arquivo}.xlsx"
CAMINHO_EXCEL = os.path.join(PASTA_ENTRADA, NOME_EXCEL)

LOG_FILE = os.path.join(LOG_FOLDER, "email_analysis.log")

inicio_execucao = datetime.now()
contador = 0

# ==================================================
# LOG
# ==================================================

def log(msg):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {msg}\n")

log("Script iniciado")

# ==================================================
# AUTENTICAÇÃO AZURE
# ==================================================

app = ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)

token = app.acquire_token_for_client(scopes=SCOPE)

if "access_token" not in token:
    log(f"Erro ao obter token: {token}")
    raise SystemExit("Falha na autenticação Azure")

headers = {
    "Authorization": f"Bearer {token['access_token']}",
    "Content-Type": "application/json"
}

log("Autenticação Azure realizada com sucesso")

# ==================================================
# FUNÇÕES AUXILIARES
# ==================================================

def limpar_texto(texto):
    return re.sub(r"\s+", " ", texto or "").strip()

def resumir_texto(texto, limite=200):
    texto = limpar_texto(texto)
    return texto[:limite] + "..." if len(texto) > limite else texto

def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).strip()
    texto = re.sub(r"\s+", " ", texto)
    return texto.lower()

# ==================================================
# LEITURA DE E-MAILS (GRAPH API)
# ==================================================

emails = []

url = (
    f"https://graph.microsoft.com/v1.0/users/{MAILBOX}/mailFolders/Inbox/messages"
    f"?$filter=receivedDateTime ge {DATA_INICIO}T00:00:00Z "
    f"and receivedDateTime le {DATA_FIM}T23:59:59Z"
    "&$top=50"
)

pagina = 1

log("Iniciando leitura da caixa de entrada")

while url:
    log(f"Lendo página {pagina}")
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    data = resp.json()

    for msg in data.get("value", []):
        contador += 1
        assunto = msg.get("subject", "")
        corpo = msg.get("body", {}).get("content", "")
        texto_total = f"{assunto} {corpo}".lower()

        if palavra_chave in texto_total:
            emails.append({
                "Título do Chamado (Assunto)": assunto,
                "Corpo do Email Resumido": resumir_texto(corpo)
            })

        if contador % 25 == 0:
            log(f"{contador} e-mails processados")

        time.sleep(0.05)

    url = data.get("@odata.nextLink")
    pagina += 1

# ==================================================
# GERAR RELATÓRIO PRINCIPAL
# ==================================================

df = pd.DataFrame(emails)
df.to_excel(CAMINHO_EXCEL, index=False)
log(f"Relatório gerado: {CAMINHO_EXCEL}")

# ==================================================
# CONSOLIDAÇÃO / RESUMO DOS CHAMADOS
# ==================================================

if not df.empty:
    df["_normalizado"] = df["Título do Chamado (Assunto)"].apply(normalizar_texto)

    resumo = (
        df[df["_normalizado"] != ""]
        .groupby("_normalizado")
        .size()
        .reset_index(name="Quantidade de Chamados")
        .sort_values(by="Quantidade de Chamados", ascending=False)
    )

    resumo["Chamado"] = resumo["_normalizado"].str.title()
    resumo = resumo[["Chamado", "Quantidade de Chamados"]]

    with pd.ExcelWriter(
        CAMINHO_EXCEL,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace"
    ) as writer:
        resumo.to_excel(writer, sheet_name="Resumo_Chamados", index=False)

    log("Resumo de chamados criado")

# ==================================================
# ORGANIZAÇÃO FINAL
# ==================================================

CAMINHO_FINAL = os.path.join(PASTA_SAIDA, NOME_EXCEL)
os.replace(CAMINHO_EXCEL, CAMINHO_FINAL)

fim_execucao = datetime.now()
duracao = fim_execucao - inicio_execucao

log(f"Processo finalizado com sucesso em {duracao}")
print("✅ Processo finalizado com sucesso")
