import os
from pathlib import Path

# ============================================================
# PATHS
# ============================================================
BASE = str(Path(__file__).resolve().parent.parent)
ENTRADA = os.path.join(BASE, "1_Agilis_Emails_Entrada")
PROCESSADOS = os.path.join(BASE, "2_Agilis_Emails_Processados")
PLANILHA_PATH = os.path.join(BASE, "program_monticelli.xlsx")
MAPAS_DIR = os.path.join(BASE, "3_Mapas_Cotacao")

# ============================================================
# EMAIL SIGNATURE
# ============================================================
EMAIL_SIGNATORY_NAME = "Ana Laura Monticelli"
EMAIL_SIGNATORY_COMPANY = "MRV Engenharia"

# ============================================================
# TRANSPORTADORAS
# ============================================================
FRACIONADO = [
    {"empresa": "GMS Transportes", "contato": "Geraldo Leite", "email": "geraldoleite@dedetransportes.com.br", "telefone": ""},
    {"empresa": "DEDE", "contato": "Vyctor Carvalho", "email": "vyctor@dedetransportes.com.br", "telefone": ""},
    {"empresa": "IRST", "contato": "Climério Neto", "email": "neto.ll@hotmail.com", "telefone": ""},
    {"empresa": "BRIX", "contato": "Raiane Fernandes", "email": "raiane.fernandes@brixcargo.com.br", "telefone": "31 98457 4885"},
    {"empresa": "HT LOG", "contato": "Lucas Nicoladelli", "email": "lucas@htlog.com.br", "telefone": ""},
    {"empresa": "VIAMAP", "contato": "Jeison Mateus", "email": "comercial@viamap.com.br", "telefone": ""},
    {"empresa": "TH TRANSPORTES", "contato": "Mauro Alves", "email": "thtransportesltda01@gmail.com", "telefone": ""},
    {"empresa": "COZAPI", "contato": "Darlan Rocha", "email": "operacional@cozapi.com.br", "telefone": ""},
]

VEICULO_COMPLETO = [
    {"empresa": "GMS Transportes", "contato": "Geraldo Leite", "email": "geraldoleite@dedetransportes.com.br", "telefone": ""},
    {"empresa": "DEDE", "contato": "Vyctor Carvalho", "email": "vyctor@dedetransportes.com.br", "telefone": ""},
    {"empresa": "HT LOG", "contato": "Lucas Nicoladelli", "email": "lucas@htlog.com.br", "telefone": ""},
    {"empresa": "VIAMAP", "contato": "Jeison Mateus", "email": "comercial@viamap.com.br", "telefone": ""},
    {"empresa": "COZAPI", "contato": "Darlan Rocha", "email": "operacional@cozapi.com.br", "telefone": ""},
]

MUDANCA = [
    {"empresa": "Montreal Mudanças", "contato": "Valter Araujo", "email": "comercial@montrealmudancas.com.br", "telefone": ""},
    {"empresa": "Fibra Mudanças", "contato": "", "email": "Contato@fibramudancas.com.br", "telefone": ""},
    {"empresa": "Fibra Mudanças", "contato": "Gentil Neto", "email": "gentil.neto@fibramudancas.com.br", "telefone": ""},
    {"empresa": "Moriway", "contato": "", "email": "comercial@moriway.com.br", "telefone": ""},
    {"empresa": "CONFIANÇA MUDANÇA", "contato": "", "email": "conf.mcz@confiancabr.com.br", "telefone": ""},
    {"empresa": "CONFIANÇA MUDANÇA", "contato": "", "email": "bocaliniwilson@yahoo.com.br", "telefone": ""},
    {"empresa": "Mudanças Tucuruvi", "contato": "Tatiana", "email": "tatiana@mudancastucuruvi.com.br", "telefone": ""},
    {"empresa": "Mudanças Trans Continental", "contato": "", "email": "mudancas@mudancastranscontinental.com.br", "telefone": ""},
    {"empresa": "Lord Mudanças", "contato": "", "email": "vendas02@lordmudancas.com.br", "telefone": ""},
    {"empresa": "Personnalite Transportes", "contato": "", "email": "comercial01@personnalitetransportes.com.br", "telefone": ""},
]

# ============================================================
# SLA
# ============================================================
SLA_RULES = {
    "NORMAL_FRETE": (2, 6, 0, 4),
    "EXPRESS_FRETE": (1, 2, 0, 1),
    "MUDANCA": (2, 5, 1, 2),
    "PESSOAS": (2, 2, 0, 1),
}

# ============================================================
# SPREADSHEET HEADERS
# ============================================================
HEADERS_CHAMADOS = [
    "Número Chamado", "Assunto", "PEP / Diagrama de Rede",
    "Data Recebimento", "Categoria",
    "SLA Step 1 (Validar)", "SLA Step 2 (Cotar)",
    "SLA Step 3 (OK Cliente)", "SLA Step 4 (Programar)",
    "Status Geral",
    "Subcategoria", "Solicitante", "Email Solicitante",
    "Telefone", "Regional", "Centro Logístico",
    "Tipo de Material", "Tipo de Veículo", "Qtd Veículos",
    "Data Coleta Solicitada", "Data Coleta Real",
    "Data Entrega",
    "Origem", "Destino",
    "Peso", "Quantidade", "Valor NF",
    "Dimensões", "Dados Complementares", "Arquivo",
    "Transportadora Vencedora", "Data Mapa Criado",
]

HEADERS_COTACOES = [
    "Número Chamado", "Transportadora", "Contato",
    "Email", "Telefone",
    "Data Resposta", "Valor Frete (R$)", "Prazo Entrega", "Observações",
]

COL_WIDTHS_CHAMADOS = {
    "Número Chamado": 16,
    "Assunto": 40,
    "PEP / Diagrama de Rede": 22,
    "Data Recebimento": 18,
    "Categoria": 18,
    "SLA Step 1 (Validar)": 18,
    "SLA Step 2 (Cotar)": 18,
    "SLA Step 3 (OK Cliente)": 20,
    "SLA Step 4 (Programar)": 20,
    "Status Geral": 14,
    "Subcategoria": 22,
    "Solicitante": 22,
    "Email Solicitante": 30,
    "Telefone": 16,
    "Regional": 16,
    "Centro Logístico": 16,
    "Tipo de Material": 22,
    "Tipo de Veículo": 18,
    "Qtd Veículos": 14,
    "Data Coleta Solicitada": 20,
    "Data Coleta Real": 18,
    "Data Entrega": 16,
    "Origem": 35,
    "Destino": 35,
    "Peso": 12,
    "Quantidade": 12,
    "Valor NF": 14,
    "Dimensões": 18,
    "Dados Complementares": 30,
    "Arquivo": 30,
    "Transportadora Vencedora": 24,
    "Data Mapa Criado": 18,
}

COL_WIDTHS_COTACOES = {
    "Número Chamado": 16,
    "Transportadora": 24,
    "Contato": 22,
    "Email": 32,
    "Telefone": 16,
    "Data Resposta": 16,
    "Valor Frete (R$)": 16,
    "Prazo Entrega": 16,
    "Observações": 35,
}

SLA_HEADER_NAMES = [
    "SLA Step 1 (Validar)",
    "SLA Step 2 (Cotar)",
    "SLA Step 3 (OK Cliente)",
    "SLA Step 4 (Programar)",
]

CHAMADOS_LEFT_ALIGN = frozenset({"Assunto", "Origem", "Destino"})
COTACOES_LEFT_ALIGN = frozenset({"Observações", "Email"})
