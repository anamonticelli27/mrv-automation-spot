import os
import re
from datetime import datetime

from frete_spot.categorization import determinar_categoria
from frete_spot.sla import calcular_slas


def extrair_campo(body, campo):
    pattern = rf"\*\s*{re.escape(campo)}\s*:\s*(.+)"
    match = re.search(pattern, body, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return ""


def extrair_numero_chamado(subject, body):
    match = re.search(r"##RE-(\d+)##", subject)
    if match:
        return match.group(1)
    match = re.search(r"N[uú]mero do Chamd?o\s*:\s*(\d+)", body, re.IGNORECASE)
    if match:
        return match.group(1)
    return ""


def extrair_assunto(body):
    match = re.search(r"\*\s*Assunto\s*:\s*(.+)", body, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return ""


def extrair_pep(body):
    match = re.search(
        r"\*\s*N[uú]mero do Elemento PEP.*?:\s*(.+)",
        body, re.IGNORECASE,
    )
    if match:
        return match.group(1).strip()
    return ""


def extrair_codigo_centro_logistico(body):
    raw = extrair_campo(body, "Centro Logístico")
    if raw:
        match = re.search(r"\b([A-Z]{2}\d{2})\b", raw.strip())
        if match:
            return match.group(1)
    return raw


def parse_received_time(received):
    try:
        return datetime(
            received.year, received.month, received.day,
            received.hour, received.minute, received.second,
        )
    except Exception:
        return datetime.now()


def parse_msg_fields(subject, body, received, filepath):
    dt_recebimento = parse_received_time(received)
    subcategoria = extrair_campo(body, "Subcategoria")
    categoria = determinar_categoria(subcategoria)
    sla1, sla2, sla3, sla4 = calcular_slas(dt_recebimento, categoria)

    return {
        "Número Chamado": extrair_numero_chamado(subject, body),
        "Assunto": extrair_assunto(body),
        "PEP": extrair_pep(body),
        "Data Recebimento": dt_recebimento,
        "Categoria": categoria,
        "SLA Step 1": sla1,
        "SLA Step 2": sla2,
        "SLA Step 3": sla3,
        "SLA Step 4": sla4,
        "Subcategoria": subcategoria,
        "Solicitante": extrair_campo(body, "Solicitante"),
        "Email Solicitante": extrair_campo(body, "E-mail"),
        "Telefone": extrair_campo(body, "Telefone/Contato"),
        "Regional": extrair_campo(body, "Regional Suprimentos"),
        "Centro Logístico": extrair_codigo_centro_logistico(body),
        "Tipo de Material": extrair_campo(body, "Tipo de Material"),
        "Tipo de Veículo": extrair_campo(body, "Tipo de veículo"),
        "Qtd Veículos": extrair_campo(body, "Quantidade de veículos necessários para carregamento"),
        "Data Coleta Solicitada": extrair_campo(body, "Data da Coleta na Origem"),
        "Data Entrega": extrair_campo(body, "Data de Entrega no Destino"),
        "Origem": extrair_campo(body, "Endereço de Origem - Completo"),
        "Destino": extrair_campo(body, "Endereço de Destino - Completo"),
        "Peso": extrair_campo(body, "Peso"),
        "Quantidade": extrair_campo(body, "Quantidade"),
        "Valor NF": extrair_campo(body, "Valor da NF"),
        "Dimensões": extrair_campo(body, "Dimensões"),
        "Dados Complementares": extrair_campo(body, "Dados Complementares"),
        "Arquivo": os.path.basename(filepath),
    }
