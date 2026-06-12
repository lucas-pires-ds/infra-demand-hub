"""
main.py — API simulada do Qualiteam
Triunfo Logística — Infraestrutura Física

Hospedagem: Render (render.com) — free tier
Endpoint principal: GET /demandas
"""

from fastapi import FastAPI, Query
from fastapi.middleware.cors import CORSMiddleware
import json
import os
from typing import Optional

app = FastAPI(
    title="Qualiteam API — Triunfo Logística (Mock)",
    description="API simulada do sistema Qualiteam para o projeto BI de Infraestrutura Física.",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["GET"],
    allow_headers=["*"],
)

# Carrega o JSON gerado pelo generate_all.py
DATA_PATH = os.path.join(os.path.dirname(__file__), "qualiteam_demandas.json")

def load_data():
    with open(DATA_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


@app.get("/")
def root():
    return {
        "sistema": "Qualiteam",
        "empresa": "Triunfo Logística",
        "versao": "1.0.0",
        "endpoints": ["/demandas", "/demandas/{id_chamado}", "/status"],
    }


@app.get("/status")
def status():
    data = load_data()
    return {
        "status": "online",
        "total_chamados": len(data),
        "sistema": "Qualiteam Mock API",
    }


@app.get("/demandas")
def get_demandas(
    status: Optional[str] = Query(None, description="Filtrar por status: Pendente, Em Andamento, Concluído"),
    prioridade: Optional[str] = Query(None, description="Filtrar por prioridade: Crítica, Alta, Média, Baixa"),
    categoria: Optional[str] = Query(None, description="Filtrar por categoria"),
):
    """
    Retorna todas as demandas de infraestrutura registradas no Qualiteam.
    Suporta filtros opcionais por status, prioridade e categoria.
    """
    data = load_data()

    if status:
        data = [d for d in data if d.get("status", "").lower() == status.lower()]
    if prioridade:
        data = [d for d in data if d.get("prioridade", "").lower() == prioridade.lower()]
    if categoria:
        data = [d for d in data if categoria.lower() in d.get("categoria", "").lower()]

    return {
        "total": len(data),
        "fonte": "Qualiteam",
        "demandas": data,
    }


@app.get("/demandas/{id_chamado}")
def get_demanda_por_id(id_chamado: str):
    """Retorna uma demanda específica pelo ID do chamado."""
    data = load_data()
    for d in data:
        if d["id_chamado"] == id_chamado:
            return d
    return {"erro": f"Chamado {id_chamado} não encontrado."}, 404
