from datetime import timedelta

from frete_spot.config import SLA_RULES


def adicionar_dias_uteis(data_inicio, dias):
    data = data_inicio
    dias_adicionados = 0
    while dias_adicionados < dias:
        data += timedelta(days=1)
        if data.weekday() < 5:
            dias_adicionados += 1
    return data


def calcular_slas(dt_recebimento, categoria):
    regras = SLA_RULES.get(categoria, (2, 6, 0, 4))
    step1_days, step2_days, step3_days, step4_days = regras

    sla_step1 = adicionar_dias_uteis(dt_recebimento, step1_days)
    sla_step2 = adicionar_dias_uteis(sla_step1, step2_days)

    if step3_days > 0:
        sla_step3 = adicionar_dias_uteis(sla_step2, step3_days)
        sla_step4 = adicionar_dias_uteis(sla_step3, step4_days)
    else:
        sla_step3 = None
        sla_step4 = adicionar_dias_uteis(sla_step2, step4_days)

    return sla_step1, sla_step2, sla_step3, sla_step4
