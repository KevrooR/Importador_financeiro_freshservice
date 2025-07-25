import customtkinter as ctk
from tkinter import filedialog, messagebox
import base64
import requests
import pandas
import datetime
import threading

base_url = "https://{url}.freshservice.com/api/v2/tickets"
api_key = "x"

def gerar_headers(api_key):
    credenciais = f"{api_key}:x"
    token = base64.b64encode(credenciais.encode()).decode()
    return {
        "Content-Type": "application/json",
        "Authorization": f"Basic {token}"
    }

def montar_chamado(linha):
    def formatar_data(data):
        if isinstance(data, (pandas.Timestamp, datetime.datetime, datetime.date)):
            return data.strftime("%Y-%m-%d")
        return data
    
    cliente = str(linha.get("CLIENTE")).strip() or "Cliente não informado"
    localizador = str(linha.get("Localizador")).strip() or "Sem localizador"
    supervisor = str(linha.get('SUPERVISOR')).strip() or "Sem supervisor"
    
    return {
        "subject": f"{cliente} – {supervisor} – {localizador}",
        "group_id": 15000761257,
        "email": str(linha.get('EMAIL SOLICITANTE')).strip() or None,
        "priority": 4,
        "status": 2,
        "description": (
        f"<div><b>Cliente:</b> {cliente}</div>"
        f"<div><b>Referência Externa:</b> {str(linha.get('Referência Externa')).strip()}</div>"
        f"<div><b>Localizador:</b> {localizador}</div>"
        f"<div><b>Pendência:</b> {str(linha.get('Pendência')).strip()}</div>"
        f"<div><b>Pedido Operação:</b> {str(linha.get('PEDIDO OPERAÇÃO')).strip()}</div>"
        f"<div><b>Email Solicitante:</b> {str(linha.get('EMAIL SOLICITANTE')).strip()}</div>"
        f"<div><b>Supervisor:</b> {supervisor}"
        f"<div><b>Self-booking:</b> {str(linha.get('Cód. Self-booking')).strip()}</div>"
        f"<div><b>Passageiro:</b> {str(linha.get('Passageiro')).strip()}</div>"
        f"<div><b>Nº Passagens:</b> {str(linha.get('Nº Passagens')).strip()}</div>"
        f"<div><b>Data Partida:</b> {formatar_data(linha['Data Partida'])}</div>"
        f"<div><b>Data Chegada:</b> {formatar_data(linha['Data Chegada'])}</div>"
        f"<div><b>Taxa Cancelamento:</b> {str(linha.get('Taxa Cancelamento')).strip()}</div>"
        f"<div><b>Desconto:</b> {str(linha.get('Desconto')).strip()}</div>"
        f"<div><b>Motivo Desconto:</b> {str(linha.get('Motivo Desconto')).strip()}</div>"
        f"<div><b>Documento:</b> {str(linha.get('Documento')).strip()}</div>"
    ),
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

def criar_ticket(dados):
    headers = gerar_headers(api_key)
    response = requests.post(base_url, headers=headers, json=dados)
    if response.status_code in [200, 201]:
        ticket_id = response.json().get("ticket", {}).get("id")
        console.insert("end", f"✅ Ticket criado - ID: {ticket_id}\n")
    else:
        console.insert("end", f"❌ Erro {response.status_code} - {response.text}\n")
    console.see("end")

def importar():
    caminho = entry_arquivo.get()
    if not caminho:
        messagebox.showerror("Erro", "Selecione um arquivo Excel.")
        return

    try:
        df = pandas.read_excel(caminho)
        df = df.where(pandas.notnull(df), None)
        df = df.convert_dtypes().astype(object)

        for _, linha in df.iterrows():
            dados = montar_chamado(linha)
            criar_ticket(dados)

        messagebox.showinfo("Finalizado", "Importação concluída.")
    except Exception as e:
        messagebox.showerror("Erro", str(e))

def iniciar_thread():
    threading.Thread(target=importar, daemon=True).start()

def selecionar_arquivo():
    caminho = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    entry_arquivo.delete(0, "end")
    entry_arquivo.insert(0, caminho)

# --- INÍCIO DA INTERFACE ---
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

root = ctk.CTk()
root.title("Importador de Chamados V2")
canvas_width = 925
canvas_height = 600
root.geometry(f"{canvas_width}x{canvas_height}")
root.resizable(False, False)

# Canvas com Formas de fundo
canvas = ctk.CTkCanvas(root, width=canvas_width, height=canvas_height, bg="white", highlightthickness=0)
canvas.pack(fill="both", expand=True)

# Triângulo decorativo
vertices = [500, 600, 800, -500, 4000, 3000]
canvas.create_polygon(vertices, outline='#E4F5FF', fill='#E4F5FF')

# Quadrado decorativo
canvas.create_rectangle(825, 0, 1000, 1000, fill="#00285C", outline="")

# --- COMPONENTES FUNCIONAIS SOBRE O CANVAS ---
label_titulo = ctk.CTkLabel(root, text="Importador - Conciliação Financeira", font=ctk.CTkFont(size=20, weight="bold"), bg_color="white", text_color="black")

label_titulo.place(x=30, y=30)

entry_arquivo = ctk.CTkEntry(root, width=400, placeholder_text="Caminho do Excel")
entry_arquivo.place(x=30, y=80)

btn_selecionar = ctk.CTkButton(root, text="Selecionar Arquivo", command=selecionar_arquivo)
btn_selecionar.place(x=440, y=80)

btn_iniciar = ctk.CTkButton(root, text="Iniciar Importação", command=iniciar_thread, fg_color="#4CAF50")
btn_iniciar.place(x=30, y=130)

console = ctk.CTkTextbox(root, width=860, height=360)
console.place(x=30, y=180)

root.mainloop()
