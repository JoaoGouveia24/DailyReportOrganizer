import pandas as pd
from openpyxl import load_workbook
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)


ascii_ART = r"""
.----.   .--.  .-..-. .-.  .-.   .----. .----..----.  .----. .----.  .---. 
| {}  \ / {} \ | || |  \ \/ /    | {}  }| {_  | {}  }/  {}  \| {}  }{_   _}
|     //  /\  \| || `--.}  {     | .-. \| {__ | .--' \      /| .-. \  | |  
`----' `-'  `-'`-'`----'`--'     `-' `-'`----'`-'     `----' `-' `-'  `-'  
 .---. .----..-. .-..----..----.   .--.  .---.  .----. .----.              
/   __}| {_  |  `| || {_  | {}  } / {} \{_   _}/  {}  \| {}  }             
\  {_ }| {__ | |\  || {__ | .-. \/  /\  \ | |  \      /| .-. \             
 `---' `----'`-' `-'`----'`-' `-'`-'  `-' `-'   `----' `-' `-'             """
print(ascii_ART)

df = pd.read_csv("report (REPLACE)/report.csv", sep=";", encoding="latin1")

#Erros a ignorar
df_Errors = pd.read_excel("Data/Data.xlsx", sheet_name="Errors")
df_Errors["ErroCode"] = df_Errors["ErroCode"].astype(str).str.strip()

#Remove os erros a ignorar
df = df[~df["HW Anomalia"].isin(df_Errors["ErroCode"])]

#Mapa de substituições
ProblemWanted = {
    "HW Anomalia": {
        "COI5001": "Encravamento no módulo de moedas",
        "CIN5001": "Encravamento no módulo de notas",
        "DSP5001": "Encravamento no módulo de notas",
        "CHQ5001": "Erro no módulo de cheques",
        "CHQ0119": "Encravamento no módulo de cheques",
        "COI5504": "Saco de moedas cheio",
        "CHQ5005": "Módulo de cheques desligado",
        "DSP5006": "Retract",
        "CIN5006": "Retract",
        "DSP5348": "Cassete tripartida cheia",
        "CIN5504": "Cassete tripartida cheia",
        "DSP5504": "Cassete tripartida cheia",
        "RPR5001": "Encravamento na impressora",
        "RPR5101": "Impressora sem papel",
        "COU5001": "Encravamento no módulo de moedas",
        "DSP5501": "Porta do cofre aberta",
        "CIN5501": "Porta do cofre aberta",
        "DSP5304": "Porta do cofre aberta",
        "CIN5010": "Máquina parada até término da operação do operador",
        "BGR5202": "Gaveta de cartões capturados cheia",
        "BGR0200": "Cartão encravado",
        "COU5006": "Erro no módulo de moedas"
    },

    "Solução": {
        "COI5001": "Verificar a máquina",
        "CIN5001": "Verificar a máquina",
        "DSP5001": "Verificar a máquina",
        "CHQ5001": "Verificar a máquina",
        "CHQ0119": "Verificar a máquina",
        "COI5504": "Esvaziar o saco",
        "CHQ5005": "Verificar a máquina",
        "DSP5006": "Esvaziar a tripartida",
        "CIN5006": "Esvaziar a tripartida",
        "DSP5348": "Esvaziar a tripartida",
        "CIN5504": "Esvaziar a tripartida",
        "DSP5504": "Esvaziar a tripartida",
        "RPR5001": "Verificar a máquina",
        "RPR5101": "Verificar a máquina",
        "COU5001": "Verificar a máquina",
        "DSP5501": "Verificar a máquina",
        "DSP5304": "Verificar a máquina",
        "CIN5501": "Verificar a máquina",
        "CIN5010": "Verificar a máquina",
        "BGR5202": "Verificar a máquina",
        "BGR0200": "Verificar a máquina",
        "COU5006": "Verificar a máquina"
    }
}

#Guarda as informações originais antes de traduzir
df["HW_Code"] = df["HW Anomalia"].astype(str).str.strip().str.upper()
df["Solução"] = df["HW_Code"].map(ProblemWanted["Solução"])
df["HW Anomalia"] = df["HW_Code"].replace(ProblemWanted["HW Anomalia"])


#Substitui os dados
df["HW Anomalia"] = df["HW Anomalia"].replace(ProblemWanted["HW Anomalia"])
df["Solução"] = df["HW Anomalia"].map(ProblemWanted["Solução"]).fillna(df["Solução"])

#Agrupa os dados por VTM
df["HW Anomalia"] = df["HW Anomalia"].astype(str).fillna("").str.strip()
df["Solução"] = df["Solução"].astype(str).fillna("").str.strip()


def concat_problemas(df_group):

    erros = df_group["HW Anomalia"].drop_duplicates()

    if len(erros) > 1:
        erros_str = " & ".join(erros)
        return f"{erros_str} - Verificar a máquina"

    erro = erros.iloc[0]
    solucao = df_group["Solução"].dropna().unique()

    if len(solucao) > 0:
        solucao_str = solucao[0]
    else:
        solucao_str = "Verificar a máquina"

    return f"{erro} - {solucao_str}"

df_Machines = pd.read_excel("Data/Data.xlsx", sheet_name="Machines")

df_Machines["Code"] = df_Machines["Code"].astype(str).str.strip()
df_Machines["Name"] = df_Machines["Name"].astype(str).str.strip()
mapa_tecnic = {}
try:
    df_Tecnic = pd.read_excel("Data/Data.xlsx", sheet_name="Tecnic", dtype=str).fillna("")
    for col in df_Tecnic.columns:
        df_Tecnic[col] = df_Tecnic[col].str.strip()

    # Se o ID estiver vazio, descobre-o pelo nome da máquina
    nome_para_code = dict(zip(df_Machines["Name"], df_Machines["Code"]))
    if "ID" not in df_Tecnic.columns:
        df_Tecnic["ID"] = ""
    df_Tecnic["ID"] = df_Tecnic.apply(
        lambda r: r["ID"] if r["ID"] else nome_para_code.get(r["Machine"], ""), axis=1)

    for codigo, grupo in df_Tecnic[df_Tecnic["ID"] != ""].groupby("ID"):
        anomalias = [a for a in grupo["Anomaly"].drop_duplicates() if a]
        owner = grupo["Owner"].iloc[0]
        mapa_tecnic[codigo] = (owner, " & ".join(anomalias))
except ValueError:
    print("Aviso: folha 'Tecnic' não encontrada, a ignorar pedidos técnicos.")

df["Código ATM"] = df["Código ATM"].astype(str).str.strip()
df = df.merge(df_Machines[["Code", "Name"]], left_on="Código ATM", right_on="Code", how="left")

df["ATM Info"] = df["Código ATM"] + " " + df["Name"]

df_agg = (
    df.groupby("ATM Info", group_keys=False)
      .apply(concat_problemas)
      .reset_index(name="Problema_Solução")
)

#Exporta os dados para um Excel limpo!!
df_agg.to_excel("CleanReport.xlsx", index=False)

import re

ficheiro_excel = "report (REPLACE)/Report_Day.xlsx"
wb = load_workbook(ficheiro_excel)
ws = wb.active  # usa a folha ativa

COL_NOME = 2     # Coluna B
COL_ACTION = 26  # Coluna Z
COL_Own = 25
COL_Tic = 24
COL_Cate = 23
LINHA_INICIO = 5

mapa_action = {}

for _, row in df_agg.iterrows():

    match = re.match(r"(\d+)", row["ATM Info"])
    if match:
        codigo = match.group(1)
        mapa_action[codigo] = row["Problema_Solução"]

linha = LINHA_INICIO
MAX_LINHAS_VAZIAS = 150
linhas_vazias = 0

while linha <= ws.max_row and linhas_vazias < MAX_LINHAS_VAZIAS:
    valor_b = ws.cell(row=linha, column=COL_NOME).value

    if not valor_b:
        linhas_vazias += 1
        linha += 1
        continue

    linhas_vazias = 0
    valor_b = str(valor_b)

    match = re.match(r"(\d+)", valor_b)
    if match:
        codigo_excel = match.group(1)
        action = mapa_action.get(codigo_excel, "")
        tecnico = mapa_tecnic.get(codigo_excel)

        if tecnico:  # prioridade ao pedido técnico
            owner, anomalia = tecnico
            ws.cell(row=linha, column=COL_Cate).value = "HW ERROR"
            ws.cell(row=linha, column=COL_Tic).value = "###"
            ws.cell(row=linha, column=COL_Own).value = owner
            ws.cell(row=linha, column=COL_ACTION).value = (
                f"{anomalia} - Enviar técnico" if anomalia else "Enviar técnico")
        elif action:  # erro automático do CSV
            ws.cell(row=linha, column=COL_Cate).value = "BRANCH"
            ws.cell(row=linha, column=COL_Tic).value = "###"
            ws.cell(row=linha, column=COL_Own).value = "NOVOBANCO"
            ws.cell(row=linha, column=COL_ACTION).value = action

    linha += 1

wb.save(ficheiro_excel)

print("Done!")
input("Pressione Enter para sair...")
