#%%
print("Importando bibliotecas...")
import sys
import numpy as np
import pandas as pd
from pathlib import Path
from unidecode import unidecode

BASE = Path(__file__).resolve().parent.parent
sys.path.append(str(BASE))

#%%
print("Importando dataframes:")
from sources.source_email import df_email
from sources.source_call import df_call
from sources.source_planner import df_planner
from sources.source_qualyteam import df_qualyteam

print("Dataframes importados com sucesso.")

# =============================================================================
# MAPEAMENTOS
# =============================================================================

MAPA_PRIORIDADES = {
    "critica":  "Crítica",
    "urgente":  "Crítica",
    "alta":     "Alta",
    "media":    "Média",
    "baixa":    "Baixa",
    "normal":   "Média",
}

MAPA_STATUS = {
    "aberto":        "Pendente",
    "em aberto":     "Pendente",
    "concluido":     "Concluído",
    "em andamento":  "Em andamento",
    "pendente":      "Pendente",
}

MAPA_CATEGORIA = {
    "Climatização / HVAC": [
        "ar condicionado", "ac ", "split", "chiller",
        "refrigeração", "temperatura", "resfriamento", "hvac",
    ],
    "Estrutura Civil": [
        "infiltração", "trinca", "viga", "laje", "cobertura",
        "piso", "afundamento", "recalque", "calçada", "cais",
        "defensa", "estrutura", "galpão",
    ],
    "Engenharia / Laudo Técnico": [
        "laudo", "inspeção técnica", "estudo estrutural",
        "capacidade de carga", "engenharia civil", "projeto e execução",
        "expansão", "rtg", "320 toneladas",
    ],
    "Segurança Física": [
        "cerca perimetral", "cerca", "ruptura",
        "rasgada", "ponto cego", "área sensível",
    ],
    "Elétrica": [
        "quadro de força", "qf-", "disjuntor", "ramal elétrico",
        "sem energia", "380v", "50a", "painel elétrico",
    ],
    "Elétrica / Iluminação": [
        "iluminação", "poste", "postes", "lâmpada", "lâmpadas",
        "luminária", "luminárias", "apagada", "apagados", "corredor escuro",
    ],
    "Elétrica / Gerador": [
        "gerador", "g2", "não partiu",
        "teste semanal", "alarme ativo no painel",
    ],
    "Elétrica / SPDA": [
        "spda", "haste quebrada", "haste coletora", "descargas atmosféricas",
    ],
    "Controle de Acesso": [
        "portão", "cancela", "trilhos", "dobradiça", "acesso",
        "entrada de veículos", "fila de caminhões", "portaria",
    ],
    "CFTV / Segurança": [
        "câmera", "camera", "cftv", "offline", "sem cobertura",
    ],
    "CFTV / Manutenção Preventiva": [
        "preventiva", "limpeza de lentes", "ajuste de foco",
        "teste de todas as câmeras", "verificação de gravação",
    ],
    "Hidráulica": [
        "vazamento", "tubulação", "água fria", "rede de água",
    ],
    "Hidráulica / Drenagem": [
        "drenagem", "canaleta", "canaletas",
        "alagamento", "chuva prevista", "obstruídas",
    ],
    "Facilities / Hidráulica": [
        "banheiro", "vaso sanitário", "entupido", "transbordando",
        "bebedouro", "refeitório", "piso molhado",
    ],
    "Sinalização / Segurança": [
        "sinalização", "faixas", "circulação",
        "não conformidade", "auditoria", "segurança",
    ],
    "Manutenção Equipamentos": [
        "empilhadeira", "freio", "sistema hidráulico", "troca de óleo",
        "filtros", "compressor", "correia transportadora", "trs-02",
    ],
    "Manutenção Preventiva": [
        "manutenção preventiva", "preventiva trimestral",
        "preventiva semestral", "verificação", "troca de óleo",
    ],
    "Equipamentos Portuários": [
        "mhc", "guindaste", "reach stacker", "cabo de aço",
        "cabo de içamento", "rtg", "içamento", "atracação", "terminal de granéis",
    ],
}

MAPA_SLA = {
    ("CFTV / Manutenção Preventiva", "Crítica"): 4,
    ("CFTV / Manutenção Preventiva", "Alta"):    24,
    ("CFTV / Manutenção Preventiva", "Média"):   48,
    ("CFTV / Manutenção Preventiva", "Baixa"):   72,

    ("CFTV / Segurança", "Crítica"): 4,
    ("CFTV / Segurança", "Alta"):    12,
    ("CFTV / Segurança", "Média"):   48,
    ("CFTV / Segurança", "Baixa"):   72,

    ("Climatização / HVAC", "Crítica"): 2,
    ("Climatização / HVAC", "Alta"):    8,
    ("Climatização / HVAC", "Média"):   48,
    ("Climatização / HVAC", "Baixa"):   48,

    ("Controle de Acesso", "Crítica"): 2,
    ("Controle de Acesso", "Alta"):    8,
    ("Controle de Acesso", "Média"):   24,
    ("Controle de Acesso", "Baixa"):   48,

    ("Elétrica", "Crítica"): 4,
    ("Elétrica", "Alta"):    24,
    ("Elétrica", "Média"):   72,
    ("Elétrica", "Baixa"):   120,

    ("Elétrica / Gerador", "Crítica"): 4,
    ("Elétrica / Gerador", "Alta"):    24,
    ("Elétrica / Gerador", "Média"):   72,
    ("Elétrica / Gerador", "Baixa"):   120,

    ("Elétrica / Iluminação", "Crítica"): 3,
    ("Elétrica / Iluminação", "Alta"):    12,
    ("Elétrica / Iluminação", "Média"):   48,
    ("Elétrica / Iluminação", "Baixa"):   48,

    ("Elétrica / SPDA", "Crítica"): 8,
    ("Elétrica / SPDA", "Alta"):    48,
    ("Elétrica / SPDA", "Média"):   96,
    ("Elétrica / SPDA", "Baixa"):   120,

    ("Engenharia / Laudo Técnico", "Crítica"): 12,
    ("Engenharia / Laudo Técnico", "Alta"):    360,
    ("Engenharia / Laudo Técnico", "Média"):   72,
    ("Engenharia / Laudo Técnico", "Baixa"):   120,

    ("Estrutura Civil", "Crítica"): 7,
    ("Estrutura Civil", "Alta"):    72,
    ("Estrutura Civil", "Média"):   96,
    ("Estrutura Civil", "Baixa"):   240,

    ("Facilities / Hidráulica", "Crítica"): 4,
    ("Facilities / Hidráulica", "Alta"):    8,
    ("Facilities / Hidráulica", "Média"):   4,
    ("Facilities / Hidráulica", "Baixa"):   24,

    ("Hidráulica", "Crítica"): 4,
    ("Hidráulica", "Alta"):    8,
    ("Hidráulica", "Média"):   24,
    ("Hidráulica", "Baixa"):   72,

    ("Hidráulica / Drenagem", "Crítica"): 4,
    ("Hidráulica / Drenagem", "Alta"):    24,
    ("Hidráulica / Drenagem", "Média"):   48,
    ("Hidráulica / Drenagem", "Baixa"):   96,

    ("Manutenção Equipamentos", "Crítica"): 6,
    ("Manutenção Equipamentos", "Alta"):    18,
    ("Manutenção Equipamentos", "Média"):   48,
    ("Manutenção Equipamentos", "Baixa"):   96,

    ("Manutenção Preventiva", "Crítica"): 12,
    ("Manutenção Preventiva", "Alta"):    24,
    ("Manutenção Preventiva", "Média"):   48,
    ("Manutenção Preventiva", "Baixa"):   96,

    ("Segurança Física", "Crítica"): 4,
    ("Segurança Física", "Alta"):    12,
    ("Segurança Física", "Média"):   48,
    ("Segurança Física", "Baixa"):   96,

    ("Sinalização / Segurança", "Crítica"): 6,
    ("Sinalização / Segurança", "Alta"):    24,
    ("Sinalização / Segurança", "Média"):   120,
    ("Sinalização / Segurança", "Baixa"):   240,

    ("Outros", "Crítica"): 8,
    ("Outros", "Alta"):    24,
    ("Outros", "Média"):   72,
    ("Outros", "Baixa"):   120,
}

# =============================================================================
# FUNÇÕES
# =============================================================================

def normalize_text(text):
    if pd.isna(text):
        return text
    return unidecode(text).lower()


def classify_category(text):

    if pd.isna(text):
        return None

    text = text.lower()

    for category, keywords in MAPA_CATEGORIA.items():

        for keyword in keywords:

            if keyword in text:
                return category

    return "Outros"


def fill_deadline(row):
    if pd.notna(row["data_limite"]):
        return row["data_limite"]
    sla_horas = MAPA_SLA.get((row["categoria"], row["prioridade"]), 24)
    return row["data_abertura"] + pd.Timedelta(hours=sla_horas)


# =============================================================================
# TRANSFORMAÇÕES POR FONTE
# =============================================================================
print("Transformando dataframes...")
#%% E-mail
df_email["hora_abertura"] = df_email["data"].dt.strftime("%H:%M")
df_email["data"] = df_email["data"].dt.date
df_email["fonte"] = "E-mail"
df_email = df_email.rename(columns={
    "data":             "data_abertura",
    "solicitante_nome": "solicitante",
    "departamento":     "setor",
    "assunto":          "descricao",
})
df_email.to_csv("df_email.csv")

#%% Call
df_call = df_call.rename(columns={
    "Data":                  "data_abertura",
    "Hora":                  "hora_abertura",
    "Quem Ligou":            "solicitante",
    "Setor":                 "setor",
    "Descrição do Problema": "descricao",
    "Urgência":              "prioridade",
    "Status":                "status",
    "Responsável":           "responsavel",
})
df_call["fonte"] = "Call"

extract = df_call["Obs"].str.extract(r'(?P<data>\d{2}/\d{2}).*?(?P<hora>\d{2}:\d{2})')
df_call["data_conclusao"] = pd.to_datetime(
    extract["data"] + "/2026 " + extract["hora"],
    format="%d/%m/%Y %H:%M",
    errors="coerce",
)
df_call["data_abertura"] = pd.to_datetime(df_call["data_abertura"]).dt.strftime("%Y-%m-%d")
df_call.to_csv("df_call.csv")

#%% Planner
df_planner = df_planner.rename(columns={
    "Data Início":    "data_abertura",
    "Data Limite":    "data_limite",
    "Nome da Tarefa": "descricao",
    "Prioridade":     "prioridade",
    "Status":         "status",
    "Atribuído a":    "responsavel",
})
df_planner["fonte"] = "Planner"
df_planner.to_csv("df_planner.csv")

#%% Qualyteam
df_qualyteam["data_abertura"] = pd.to_datetime(df_qualyteam["data_abertura"])
df_qualyteam["hora_abertura"] = df_qualyteam["data_abertura"].dt.strftime("%H:%M")
df_qualyteam["data_abertura"] = df_qualyteam["data_abertura"].dt.date
df_qualyteam = df_qualyteam.rename(columns={
    "prazo_conclusao":   "data_limite",
    "setor_solicitante": "setor",
    "responsavel_infra": "responsavel",
})
df_qualyteam["fonte"] = "Qualyteam"
df_qualyteam.to_csv("df_qualyteam.csv")

# =============================================================================
# CONSOLIDAÇÃO
# =============================================================================
print("Consolidando em um unico dataframe.")
#%%
COLUNAS_TEMPLATE = [
    "data_abertura", "hora_abertura", "data_limite", "data_conclusao",
    "solicitante", "setor", "categoria", "descricao",
    "prioridade", "responsavel", "status", "fonte",
]

df_template = pd.DataFrame(columns=COLUNAS_TEMPLATE)

df = pd.concat([df_template, df_email, df_call, df_qualyteam]).reset_index(drop=True)

COLUNAS_FINAIS = [
    "data_abertura", "data_limite", "data_conclusao",
    "solicitante", "setor", "categoria", "descricao",
    "prioridade", "responsavel", "status", "fonte",
]

df_final = df[COLUNAS_FINAIS].copy()

# =============================================================================
# NORMALIZAÇÃO
# =============================================================================

#%%
df_final["prioridade"] = df_final["prioridade"].apply(normalize_text).map(MAPA_PRIORIDADES)
df_final["status"]     = df_final["status"].apply(normalize_text).map(MAPA_STATUS)

df_final["data_abertura"] = pd.to_datetime(
    df_final["data_abertura"].astype(str) + " " + df["hora_abertura"],
    format="%Y-%m-%d %H:%M",
)

df_final["data_limite"] = pd.to_datetime(df_final["data_limite"], errors="coerce")

df_final["data_conclusao"] = pd.to_datetime(df_final["data_conclusao"], errors="coerce")

df_final["categoria"]   = df_final["descricao"].apply(classify_category)
df_final["data_limite"] = df_final.apply(fill_deadline, axis=1)

df_final["responsavel"] = (
    df_final["responsavel"]
    .fillna("Não atribuído")
    .replace("nao atribuido", "Não atribuído")
)

# =============================================================================
# STATUS SLA
# =============================================================================

#%%
condicoes = [
    (df_final["status"] == "Concluído")    & (df_final["data_conclusao"] <= df_final["data_limite"]),
    (df_final["status"] == "Concluído")    & (df_final["data_conclusao"] >  df_final["data_limite"]),
    (df_final["status"] == "Concluído")    & (df_final["data_conclusao"].isna() == True),
    (df_final["status"] == "Pendente")     & (df_final["data_conclusao"] <= df_final["data_limite"]),
    (df_final["status"] == "Pendente")     & (df_final["data_conclusao"] >  df_final["data_limite"]),
    (df_final["status"] == "Em andamento") & (df_final["data_conclusao"] <= df_final["data_limite"]),
    (df_final["status"] == "Em andamento") & (df_final["data_conclusao"] >  df_final["data_limite"]),
]

status_sla = [
    "Concluído no prazo", "Concluído com atraso", "Não preenchido",
    "No prazo", "Vencido",
    "No prazo", "Vencido",
]

df_final["status_SLA"] = np.select(condlist=condicoes, choicelist=status_sla, default=None)

print("Disponibilizando tabela transformada em csv...")

df_final.to_csv("demandas_consolidadas.csv", index=False)

print("Tabela consolidada salva com sucesso.")