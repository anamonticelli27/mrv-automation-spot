from frete_spot.categorization import determinar_lista_transportadoras
from frete_spot.email.templates import gerar_corpo_frete, gerar_corpo_mudanca


def open_shared_msg(filepath):
    import win32com.client

    outlook_ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    return outlook_ns.OpenSharedItem(filepath)


def criar_rascunho_outlook(dados, categoria):
    import win32com.client

    chamado = dados.get("Número Chamado", "SEM_NUMERO")
    lista = determinar_lista_transportadoras(categoria, dados.get("Tipo de Veículo", ""))

    if categoria == "MUDANCA":
        corpo = gerar_corpo_mudanca(dados)
    elif categoria == "EXPRESS_FRETE":
        corpo = gerar_corpo_frete(dados, "Frete Expresso")
    elif categoria == "PESSOAS":
        corpo = gerar_corpo_frete(dados, "Frete Expresso - Pessoas")
    else:
        corpo = gerar_corpo_frete(dados, dados.get("Tipo de Veículo", "Frete Normal"))

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.Subject = f"COTAÇÃO SERVIÇO DE TRANSPORTE DE MATERIAL - {chamado}"
        mail.Body = corpo
        mail.To = ""

        for transportadora in lista:
            recipient = mail.Recipients.Add(transportadora["email"])
            recipient.Type = 3  # olBCC

        mail.Recipients.ResolveAll()
        mail.Save()

        print(f"   📧 Rascunho criado no Outlook para chamado {chamado} ({len(lista)} transportadoras em BCC)")

    except Exception as e:
        print(f"   ❌ Erro ao criar rascunho no Outlook: {e}")
