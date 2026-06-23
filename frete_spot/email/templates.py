from frete_spot.config import EMAIL_SIGNATORY_COMPANY, EMAIL_SIGNATORY_NAME


def gerar_corpo_frete(dados, modalidade):
    return f"""Prezado Transportador,

Boa tarde!

Solicito cotação de frete conforme dados abaixo:

• Chamado Agilis : {dados.get('Número Chamado', '')}
• Modalidade     : {modalidade}
• Tipo Material  : {dados.get('Tipo de Material', '')}
• Origem         : {dados.get('Origem', '')}
• Destino        : {dados.get('Destino', '')}
• Data Coleta    : {dados.get('Data Coleta Solicitada', '')}
• Data Entrega   : {dados.get('Data Entrega', '')}
• Peso           : {dados.get('Peso', '')}
• Quantidade     : {dados.get('Quantidade', '')}
• Dimensões      : {dados.get('Dimensões', '')}
• Valor da NF    : {dados.get('Valor NF', '')}
• Observações    : {dados.get('Dados Complementares', '')}

Por favor, retornar com:
✔ Valor do frete
✔ Prazo de entrega
✔ Código interno (se houver)

Aguardo retorno.

Atenciosamente,
{EMAIL_SIGNATORY_NAME}
{EMAIL_SIGNATORY_COMPANY}"""


def gerar_corpo_mudanca(dados):
    return f"""Prezado Transportador,

Boa tarde!

Solicito cotação de mudança conforme dados abaixo:

• Chamado Agilis : {dados.get('Número Chamado', '')}
• Solicitante    : {dados.get('Solicitante', '')}
• Origem         : {dados.get('Origem', '')}
• Destino        : {dados.get('Destino', '')}
• Data Coleta    : {dados.get('Data Coleta Solicitada', '')}
• Data Entrega   : {dados.get('Data Entrega', '')}
• Observações    : {dados.get('Dados Complementares', '')}

Por favor, retornar com:
✔ Valor do frete
✔ Prazo de entrega
✔ Código interno (se houver)

Aguardo retorno.

Atenciosamente,
{EMAIL_SIGNATORY_NAME}
{EMAIL_SIGNATORY_COMPANY}"""
