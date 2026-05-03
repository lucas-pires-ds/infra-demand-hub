#%%
import pandas as pd
import os
from pathlib import Path

BASE = Path().absolute()


#%%

from sources.source_email import df_email
from sources.source_call import df_call
from sources.source_planner import df_planner 
from sources.source_qualyteam import df_qualyteam

#%%

print("df_email: \n", df_email.columns)
print("\ndf_call: \n", df_call.columns)
print("\ndf_planner: \n", df_planner.columns)
print("\ndf_qualyteam: \n", df_qualyteam.columns)


df_template = pd.DataFrame(columns=["data_abertura", "hora_abertura","data_limite", "solicitante", "setor", "descricao", "prioridade", "status", "fonte"])
display(df_template)
#%%

df_email["hora_abertura"] = df_email["data"].dt.strftime("%H:%M")
df_email["data"] = df_email["data"].dt.date
df_email["fonte"] = "E-mail"
df_email = df_email.rename(columns={
    "data" : "data_abertura",
    "solicitante_nome" : "solicitante",
    "departamento" : "setor",
    "assunto" : "descricao",
                                    }
                            )

df_call = df_call.rename(columns={
    "Data" : "data_abertura",
    "Hora" : "hora_abertura",
    "Quem Ligou" : "solicitante",
    "Setor" : "setor",
    "Descrição do Problema" : "descricao",
    "Urgência" : "prioridade",
    "Status" : "status"
})
df_call["fonte"] = "Call"

df_planner = df_planner.rename(columns={
    "Data Início" : "data_abertura",
    "Data Limite" : "data_limite",
    "Nome da Tarefa" : "descricao",
    "Prioridade" : "prioridade",
    "Status" : "status"
    })
df_planner["fonte"] = "Planner"

df_qualyteam["data_abertura"] = pd.to_datetime(df_qualyteam["data_abertura"])
df_qualyteam["hora_abertura"] = df_qualyteam["data_abertura"].dt.strftime("%H:%M")
df_qualyteam["data_abertura"] = df_qualyteam["data_abertura"].dt.date

df_qualyteam = df_qualyteam.rename(columns={
    "prazo_conclusao" : "data_limite",
    "setor_solicitante" : "setor"})
df_qualyteam["fonte"] = "Qualyteam"


#%%
df = pd.concat([df_template, df_email, df_call, df_qualyteam]).reset_index()
df_final = df[df_template.columns]





