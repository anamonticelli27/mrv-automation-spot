import gc
import os
import shutil
import time

from frete_spot.categorization import determinar_lista_transportadoras
from frete_spot import config
from frete_spot.excel.format import formatar_planilha
from frete_spot.excel.workbook import fmt_data, inicializar_planilha, salvar_cotacoes, salvar_na_planilha
from frete_spot.extractors import parse_msg_fields


def processar_msg(filepath):
    from frete_spot.email.outlook import criar_rascunho_outlook, open_shared_msg

    msg = open_shared_msg(filepath)

    subject = msg.Subject or ""
    body = msg.Body or ""
    received = msg.ReceivedTime

    dados = parse_msg_fields(subject, body, received, filepath)
    categoria = dados["Categoria"]
    sla1, sla2, sla3, sla4 = (
        dados["SLA Step 1"],
        dados["SLA Step 2"],
        dados["SLA Step 3"],
        dados["SLA Step 4"],
    )

    print(f"\n📨 Chamado  : {dados['Número Chamado']}")
    print(f"   Assunto   : {dados['Assunto']}")
    print(f"   PEP       : {dados['PEP']}")
    print(f"   Categoria : {categoria}")
    print(f"   SLA Step1 : {fmt_data(sla1)}")
    print(f"   SLA Step2 : {fmt_data(sla2)}")
    if sla3:
        print(f"   SLA Step3 : {fmt_data(sla3)}")
    print(f"   SLA Step4 : {fmt_data(sla4)}")

    lista_transportadoras = determinar_lista_transportadoras(categoria, dados.get("Tipo de Veículo", ""))

    salvar_na_planilha(dados)
    salvar_cotacoes(dados, lista_transportadoras)
    criar_rascunho_outlook(dados, categoria)

    msg.Close(0)
    del msg
    gc.collect()
    time.sleep(1)

    destino = os.path.join(config.PROCESSADOS, os.path.basename(filepath))
    try:
        shutil.move(filepath, destino)
        print("   📁 Movido para: 2_Agilis_Emails_Processados")
    except Exception as e:
        print(f"   ⚠️  Não foi possível mover o arquivo: {e}")


def main():
    inicializar_planilha()
    arquivos = [f for f in os.listdir(config.ENTRADA) if f.lower().endswith(".msg")]

    if not arquivos:
        print("📭 Nenhum email em 1_Agilis_Emails_Entrada.")
        formatar_planilha()
        return

    print(f"📬 {len(arquivos)} email(s) encontrado(s).\n")
    for arquivo in arquivos:
        filepath = os.path.join(config.ENTRADA, arquivo)
        try:
            processar_msg(filepath)
        except Exception as e:
            print(f"❌ Erro ao processar {arquivo}: {e}")

    formatar_planilha()
    print("\n✅ Concluído!")
