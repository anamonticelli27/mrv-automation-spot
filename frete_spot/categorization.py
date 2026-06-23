from frete_spot.config import FRACIONADO, MUDANCA, VEICULO_COMPLETO


def determinar_categoria(subcategoria):
    sub = subcategoria.lower().strip()
    if "pessoas" in sub:
        return "PESSOAS"
    if "expresso" in sub or "expres" in sub:
        return "EXPRESS_FRETE"
    if "mudan" in sub:
        return "MUDANCA"
    return "NORMAL_FRETE"


def determinar_lista_transportadoras(categoria, tipo_veiculo):
    vei = tipo_veiculo.lower().strip()
    if categoria == "MUDANCA":
        return MUDANCA
    if "fracionado" in vei:
        return FRACIONADO
    return VEICULO_COMPLETO
