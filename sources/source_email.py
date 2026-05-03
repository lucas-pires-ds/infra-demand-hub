#%%
import pandas as pd
from pathlib import Path
from email import policy
from email.parser import BytesParser

BASE = Path().absolute()
input_dir = BASE / "mock_data/output"

pd.set_option('display.max_colwidth', None)


emails_dir = Path(input_dir / "emails_eml")


MAPA_CATEGORIAS_ITEM = {
    "câmera 14": "CFTV",
    "ar condicionado": "Refrigeração",
    "infiltração": "Civil/Infraestrutura",
    "portão": "Acessos e Portaria",
    "postes": "Iluminação",
    "bebedouro": "Facilidades",
    "gerador": "Equipamentos Críticos",
    "banheiro": "Hidráulica/Civil",
    "cerca perimetral": "Segurança Patrimonial",
    "faixas de sinalização": "Sinalização/Segurança",
    "refrigeração": "Refrigeração",
    "luminárias": "Iluminação",
    "viga": "Estrutural",
    "empilhadeira": "Frota/Equipamentos",
    "piso solto": "Civil/Infraestrutura",
    "laudo técnico": "Documentação/Engenharia",
    "canaletas de drenagem": "Hidráulica/Infraestrutura",
    "haste coletora": "SPDA (Para-raios)",
    "split": "Refrigeração",
    "cancela eletrônica": "Acessos e Portaria"
}


MAPA_LOCAIS = {
    "cais principal": "Cais Principal",
    "sala de reuniões do 2° andar": "Sala Reuniões (2° Piso)",
    "galpão 3": "Galpão 03",
    "galpão 1": "Galpão 01",
    "pátio c": "Pátio C",
    "refeitório": "Refeitório",
    "térreo": "Térreo",
    "pátio norte": "Pátio Norte",
    "pátio operacional": "Pátio Operacional",
    "administrativo no 1° andar": "ADM (1° Piso)",
    "portaria principal": "Portaria Principal",
    "pátio sul": "Pátio Sul",
    "galpão 4": "Galpão 04",
    "galpão 2": "Galpão 02",
    "sala de operações do térreo": "Sala Operações (Térreo)",
    "portaria 2": "Portaria 2"
}

MAPA_DEPARTAMENTOS = {
    "Jorge Salles": "Segurança Patrimonial",
    "Paula Drummond": "Administrativo",
    "Sérgio Pimentel": "Operações Portuárias",
    "Fábio Gomes": "Manutenção",
    "Rogério": "Operações Portuárias",
    "Rogério Andrade": "Operações Portuárias",
    "Beatriz Carvalho": "RH",
    "Antônio Ferreira": "Manutenção",
    "Claudia": "Operacional",
    "Cristiane Barros": "Compliance",
    "Renata": "Financeiro",
    "Marcos Vinicius Teles": "Planejamento"
}



MAPA_REGRAS_PRIORIDADE = {
    "IMEDIATA": "Crítica",
    "Urgente": "Crítica",
    "acidente grave": "Crítica",
    "risco real": "Crítica",
    "inaceitável": "Crítica",
    "não conformidade com a NR-12": "Crítica",
    "insustentável": "Crítica",
    "interditar a área": "Crítica",
    "grave": "Crítica",
    "atrasar a operação": "Alta",
    "sem monitoramento": "Alta",
    "sem backup": "Alta",
    "piorou bastante": "Alta",
    "priorizar": "Alta",
    "reparo urgente": "Alta",
    "estamos com frota reduzida": "Alta",
    "alagamento é certo": "Alta",
    "Já reportei" : "Alta",
    "quanto antes" : "Alta",
    "resolver hoje": "Alta",
    "segunda vez esse mês": "Média",
    "3ª vez e ainda sem solução": "Média",
    "Prazo apertado": "Média",
    "prevista para o mês que vem": "Média",
    "antes de sexta": "Média",
    "se puder resolver essa semana": "Baixa",
    "talvez precise de substituição": "Baixa",
    "prevendo a instalação": "Baixa"
}

registros_emails = []
for i, caminho_email in enumerate(emails_dir.iterdir()):
    with open(caminho_email, "rb") as f:
        msg = BytesParser(policy=policy.default).parse(f)
    
    dados_extraidos = {
        "data" : pd.to_datetime(msg["Date"]),
        "solicitante_nome": msg["From"].partition("<")[0].replace('"', ''),
        "solicitante_email" :  msg["From"].partition("<")[2].replace('"', '').replace(">", ""),
        "local" : None,
        "item" : None,
        "departamento" : None,
        "prioridade" : None,
        "assunto" : msg["Subject"],
        "responsavel" : "nao atribuido",
        "status" : "aberto"
        }

    for part in msg.walk():
        
        if part.get_content_type() == 'text/plain':
            texto = part.get_payload(decode=True).decode().lower()
             
            for departamento in MAPA_DEPARTAMENTOS:
                if departamento.lower() in texto:
                    dados_extraidos["departamento"] = MAPA_DEPARTAMENTOS[departamento]
                    break
                
            for item in MAPA_CATEGORIAS_ITEM:
                if item.lower() in texto:
                    dados_extraidos["item"] = MAPA_CATEGORIAS_ITEM[item]
                    break
                
            for local in MAPA_LOCAIS:
                if local.lower() in texto:
                    dados_extraidos["local"] = MAPA_LOCAIS[local]
                    break
                
            for prioridade in MAPA_REGRAS_PRIORIDADE:
                if prioridade.lower() in texto:
                    dados_extraidos["prioridade"] = MAPA_REGRAS_PRIORIDADE[prioridade]
                    break
                
            
    registros_emails.append(dados_extraidos)
    


print("definindo df_email...")
df_email = pd.DataFrame(registros_emails)