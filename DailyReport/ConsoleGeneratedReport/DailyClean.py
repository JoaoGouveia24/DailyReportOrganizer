import pandas as pd

# === 1. Ler CSV original ===
df = pd.read_csv("report (REPLACE)/report.csv", sep=";", encoding="latin1")

Erros_Ignorar = ["CIN5505",
                 "COU6322", #Hoppers vazios
                 "DSP5322",
                 "CHQ0004",
                 "EID5253", #Cartão de cidadão inválido
                 "DSP5318", #Cassetes vazias
                 "COU0306", #Valor de moedas não dispensado
                 "COI0014", #Erro de hardware nas moedas
                 "CHQ0102", #Nenhuma verificação feita no módulo
                 "0000000", #Invalid fields
                 "CHQ0129", #Tamanho de verificação invalido
                 "0000048", #Timeout autenticação
                 "RPR5100",
                 "CIN0014",
                 "EID5254",
                 "CHQ0014",
                 "EID5255",
                 "BGR5201",
                 "RPR0123",
                 "BGR5006",
                 "RPR0122",
                 "COI5006",
                 "0100002",
                 "0001011",
                 "0000099",
                 "EID5252", #Não foi inserido cartão algum
                 "CHQ0013",
                 "CHQ0053",
                 "DSP5307",
                 "DSP0001",
                 "CHQ0041"
                 ]

df = df[~df["HW Anomalia"].isin(Erros_Ignorar)]

# === 2. Mapa de substituições ===
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
        "EID5250": "Leitor cartão cidadão não conectado",
        "BGR5202": "Gaveta de cartões capturados cheia",
        "BGR0200": "Cartão encravado"
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
        "EID5250": "Verificar a máquina",
        "BGR5202": "Verificar a máquina",
        "BGR0200": "Verificar a máquina"
    }
}

# Guardar o código original antes de traduzir
df["HW_Code"] = df["HW Anomalia"].astype(str).str.strip().str.upper()
df["Solução"] = df["HW_Code"].map(ProblemWanted["Solução"])
df["HW Anomalia"] = df["HW_Code"].replace(ProblemWanted["HW Anomalia"])


# === 3. Substituir ===
df["HW Anomalia"] = df["HW Anomalia"].replace(ProblemWanted["HW Anomalia"])
df["Solução"] = df["HW Anomalia"].map(ProblemWanted["Solução"]).fillna(df["Solução"])

# === 4. Agrupar por ATM ===
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


df_agg = df.groupby("Código ATM").apply(concat_problemas).reset_index(name="Problema_Solução")

# === 5. Exportar ===
df_agg.to_excel("CleanReport.xlsx", index=False)

print("✅ Arquivo 'CleanReport.xlsx' gerado com sucesso!")
