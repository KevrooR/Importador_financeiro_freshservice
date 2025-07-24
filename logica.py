import base64
import requests
import pandas
import datetime

base_url = "https://copasturhelpdesk.freshservice.com/api/v2/tickets"
solicitante_url = "https://copasturhelpdesk.freshservice.com/api/v2/requesters?email="
api_key = "XXXXXXXXXXXXX"
arquivo_excel = "C:/Users/kevin.araujo/OneDrive - copastur.com.br/Área de Trabalho/1051 CANCELADOS_teste.xlsx"

df = pandas.read_excel(arquivo_excel)
df = df.where(pandas.notnull(df), None)
df = df.convert_dtypes().astype(object)

print(df.columns)
    
def gerar_headers(api_key):
    credenciais = f"{api_key}:x"
    token = base64.b64encode(credenciais.encode()).decode()
    return {
        "Content-Type":"application/json",
        "Authorization":f"Basic {token}"
    }

def montar_chamado(linha):
    from pandas import notnull
    
    linha = linha.where(pandas.notnull(linha), None)
    
    def formatar_data(data):
        if isinstance(data, (pandas.Timestamp, datetime.datetime, datetime.date)):        return data.strftime("%Y-%m-%d")
        return data
    
    chamado = {
    "subject": "Teste, Criação em massa",
    "group_id": 15000761257,
    "email": str(linha.get('EMAIL SOLICITANTE')).strip() or None,
    "priority": 4,
    "status": 2,
    "description": f"<div>TESTE</div><div>{linha['CLIENTE']}</div>",
    "custom_fields": {
        "tipo": "Financeiro teste",
        "cd_qp": str(linha.get('Cód. QP')).strip() or None,
        "referncia_externa": str(linha.get('Referência Externa')).strip() or None,
        "localizador": str(linha.get('Localizador')).strip() or None,
        "pendncia": str(linha.get('Pendência')).strip() or None,
        "pedido_operao": str(linha.get('PEDIDO OPERAÇÃO')).strip() or None,
        "email_solicitante": str(linha.get('EMAIL SOLICITANTE')).strip() or None,
        "supervisor": str(linha.get('SUPERVISOR')).strip() or None,
        "cd_selfbooking": str(linha.get('Cód. Self-booking')).strip() or None,
        "n_pasagens": str(linha.get('Nº Passagens')).strip() or None,
        "data_partida": formatar_data(linha["Data Partida"]),
        "data_chegada": formatar_data(linha["Data Chegada"]),
        "taxa_cancelamento": str(linha.get('Taxa Cancelamento')).strip() or None,
        "desconto": str(linha.get('Desconto')).strip() or None,
        "motivo_desconto": str(linha.get('Motivo Desconto')).strip() or None,
        "documento": str(linha.get('Documento')).strip() or None

    }
}
    return chamado
    
def criar_ticket(dados):
    headers = gerar_headers(api_key)
    response = requests.post(base_url, headers=headers, json=dados)
    
    if response.status_code in [200, 201]:
        ticket_id = response.json().get("ticket", {}).get("id")
        print(f"Ticket criado com sucesso - ID: {ticket_id}")
    else:
        print(f"Erro de criação: {response.status_code} - {response.text}")
        
for index, linha in df.iterrows():
    dados = montar_chamado(linha)
    criar_ticket(dados)
