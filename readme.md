# infra-demand-hub

Dashboard de centralização de demandas do setor de infraestrutura física. O projeto consolida chamados vindos de quatro fontes distintas em um único painel, eliminando a necessidade de monitorar sistemas separados para garantir prazos e requisitos técnicos.

---

## Problema

O gestor de infraestrutura recebe demandas por quatro canais diferentes e precisa acessar cada um manualmente para verificar status, prazos e responsáveis. Demandas de infra têm SLAs curtos — algumas com prazo de horas — o que torna o monitoramento fragmentado um risco operacional real.

## Solução

Pipeline de dados automatizado que extrai, transforma e consolida as quatro fontes em um banco central, alimentando um dashboard Power BI atualizado a cada hora sem depender de intervenção manual.

---

## Arquitetura

```
[Qualiteam]     [Planner/Teams]     [E-mail]     [Call/Telefone]
   API REST          CSV               .eml            XLSX
      |                |                 |                |
      └────────────────┴─────────────────┴────────────────┘
                               |
                     [ETL — Python Scripts]
                               |
                     [PostgreSQL — Supabase]
                               |
                    [Power BI Service — Dashboard]
```

**Orquestração:** GitHub Actions (cron a cada 1 hora)
**Banco central:** PostgreSQL no Supabase (free tier)
**API simulada:** FastAPI hospedada no Render (free tier)
**Visualização:** Power BI Service

---

## Fontes de Dados

| Fonte | Sistema | Formato | Hospedagem |
|---|---|---|---|
| Qualiteam | Sistema de gestão de qualidade | JSON via API REST | FastAPI no Render |
| Planner | Microsoft Teams / Planner | CSV exportado | Google Drive |
| E-mail | Caixa de entrada do setor | Arquivos `.eml` | Pasta local / simulada |
| Call | Registro telefônico manual | Excel `.xlsx` | Google Drive |

Cada fonte possui inconsistências propositais que simulam a realidade: campos nulos, formatos de data divergentes, status com variações de escrita e prioridades sem padronização. O ETL é responsável por tratar tudo isso antes de gravar no banco central.

---

## Estrutura do Repositório

```
infra-demand-hub/
├── .github/
│   └── workflows/
│       └── etl_pipeline.yml        # scheduler GitHub Actions (cron horário)
├── api_qualiteam/
│   ├── main.py                     # FastAPI — endpoint /demandas
│   ├── requirements.txt
│   └── qualiteam_demandas.json     # dados servidos pela API
├── mock_data/
│   ├── generate_all.py             # gerador de todos os dados fictícios
│   └── output/
│       ├── qualiteam_demandas.json
│       ├── planner_tarefas.csv
│       ├── call_chamados.xlsx
│       └── emails_eml/             # 20 arquivos .eml
├── sources/
│   ├── source_qualiteam.py         # consome a API REST
│   ├── source_planner.py           # lê o CSV do Google Drive
│   ├── source_email.py             # faz parse dos arquivos .eml
│   └── source_call.py              # lê o XLSX
├── transform/
│   └── consolidate.py              # padroniza e une as quatro fontes
├── load/
│   └── update_database.py          # grava no PostgreSQL (Supabase)
├── dashboard/
│   └── infra_demand_hub.pbix       # arquivo Power BI
├── requirements.txt
└── README.md
```

---

## Pré-requisitos

- Python 3.11+
- Conta no [Supabase](https://supabase.com) (gratuito)
- Conta no [Render](https://render.com) (gratuito)
- Conta no GitHub
- Power BI Desktop + Power BI Service (conta Microsoft)

---

## Configuração

### 1. Variáveis de ambiente

Crie um arquivo `.env` na raiz do projeto:

```env
SUPABASE_URL=https://xxxx.supabase.co
SUPABASE_KEY=your_anon_key
API_QUALITEAM_URL=https://sua-api.onrender.com
PLANNER_CSV_URL=https://drive.google.com/uc?id=SEU_FILE_ID
CALL_XLSX_URL=https://drive.google.com/uc?id=SEU_FILE_ID
EMAIL_DIR=./mock_data/output/emails_eml
```

### 2. Instalar dependências

```bash
pip install -r requirements.txt
```

### 3. Subir a API no Render

1. Faça fork ou push do projeto no GitHub
2. No Render, crie um novo **Web Service** apontando para a pasta `api_qualiteam/`
3. Start command: `uvicorn main:app --host 0.0.0.0 --port 10000`
4. Copie a URL gerada e cole em `API_QUALITEAM_URL` no `.env`

### 4. Configurar os Secrets no GitHub

No repositório, vá em **Settings → Secrets and variables → Actions** e adicione:

- `SUPABASE_URL`
- `SUPABASE_KEY`
- `API_QUALITEAM_URL`
- `PLANNER_CSV_URL`
- `CALL_XLSX_URL`

### 5. Rodar o ETL manualmente (teste)

```bash
python transform/consolidate.py
```

---

## Agendamento Automático

O arquivo `.github/workflows/etl_pipeline.yml` configura o GitHub Actions para rodar o pipeline a cada hora. Nenhuma máquina local precisa estar ligada.

```yaml
on:
  schedule:
    - cron: '0 * * * *'
```

---

## Dashboard

O arquivo `infra_demand_hub.pbix` se conecta ao PostgreSQL no Supabase e apresenta:

- Painel de alertas com demandas vencidas e próximas do prazo
- Visão Kanban por status (Pendente / Em Andamento / Concluído)
- Volume de chamados por fonte e por categoria
- SLA médio de resolução por responsável
- Filtros por período, prioridade, setor e fonte
- Timestamp de última atualização

---

## Dados Fictícios

Os dados foram gerados pelo script `mock_data/generate_all.py` e simulam um terminal portuário com operação 24 horas. As demandas cobrem desde laudos de capacidade de solo para guindastes de centenas de toneladas até manutenções simples de facilities.

Para regenerar os dados:

```bash
cd mock_data
python generate_all.py
```

---

## Stack

| Camada | Tecnologia |
|---|---|
| Linguagem | Python 3.11 |
| API simulada | FastAPI + Uvicorn |
| Orquestração | GitHub Actions |
| Banco de dados | PostgreSQL (Supabase) |
| Visualização | Power BI Service |
| Hospedagem API | Render |

---

## Sobre o Projeto

Este projeto foi desenvolvido como portfólio para a posição de **Analista de Infraestrutura** com foco em Business Intelligence. O objetivo é demonstrar a capacidade de estruturar uma solução completa de centralização de dados — desde a ingestão de múltiplas fontes heterogêneas até a entrega de um dashboard operacional com atualização automatizada.
