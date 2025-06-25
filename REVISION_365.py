# -*- coding: utf-8 -*-
"""
Created on Wed Jun 25 12:14:30 2025

@author: 67397190
"""

import tkinter as tk
from tkinter import messagebox, scrolledtext, simpledialog
from msal import PublicClientApplication
import requests
from bs4 import BeautifulSoup
import re
import pandas as pd
import os
import base64
import tempfile
import subprocess
import datetime
from fuzzywuzzy import fuzz, process
    
# === CONFIGURACIÓN FIJA ===
config = {}

with open("CONFIG365.txt", "r") as file:
    for line in file:
        if "=" in line:
            key, value = line.strip().split("=", 1)
            config[key.strip()] = value.strip()

# Asignar a variables
client_id = config.get("client_id")
tenant_id = config.get("tenant_id")
carpeta_id = config.get("carpeta_id")
carpeta_descarga = 'C:/FACTURAS'
os.makedirs(carpeta_descarga, exist_ok=True)
    
# === GUI ===
root = tk.Tk()
root.title("Revisión de Correos")
frame_top = tk.Frame(root)
frame_top.pack(pady=5)
tk.Label(frame_top, text="Número de correos a revisar:").pack(side=tk.LEFT)
entry_n = tk.Entry(frame_top, width=5)
entry_n.insert(0, "5")
entry_n.pack(side=tk.LEFT)
btn_iniciar = tk.Button(frame_top, text="Iniciar Revisión")
btn_iniciar.pack(side=tk.LEFT, padx=10)
frame_main = tk.Frame(root)
frame_main.pack()
text_campos = scrolledtext.ScrolledText(frame_main, width=60, height=30)
text_campos.pack(side=tk.LEFT, padx=5)
text_cuerpo = scrolledtext.ScrolledText(frame_main, width=80, height=30)
text_cuerpo.pack(side=tk.LEFT, padx=5)
    
avanzar_event = tk.BooleanVar()
usuario = os.getlogin()
    
def extraer_dominio(correo):
    match_email = re.search(r'<([^<>]+@[^<>]+)>', str(correo))
    if match_email:
        correo = match_email.group(1)
    else:
        match_email = re.search(r'\b[^\s<>]+@[^\s<>]+\b', str(correo))
        if match_email:
            correo = match_email.group(0)
        else:
            return ""
        match_dominio = re.search(r'@([\w\.-]+)', correo)
            if match_dominio:
        dominio = match_dominio.group(1)
            dominio = re.sub(r'\.(com|es|net|org|edu|gov|info|biz|co|ar|mx|cl|fr|de|uk|it|pt|br|us|ca|au|ch|nl|be|no|se|fi|dk|pl|cz|ru|cn|jp|kr|in)$', '', dominio, flags=re.IGNORECASE)
            return dominio.lower()
            return ""

def extraer_bloques_encabezado(texto):
    bloques = re.findall(r'(De:.*?Asunto:.*?)(?=(?:\nDe:)|\Z)', texto, re.DOTALL | re.IGNORECASE)
    return bloques

def extraer_campos_desde_bloque(bloque):
    campos = {'De': '', 'Enviado': '', 'Para': '', 'Cc': '', 'Asunto': ''}
    patrones = {
        'De': r'De:\s*(.*)',
        'Enviado': r'Enviado:\s*(.*)',
        'Para': r'Para:\s*(.*)',
        'Cc': r'Cc:\s*(.*)',
        'Asunto': r'Asunto:\s*(.*)'
    }
    for campo, patron in patrones.items():
match = re.search(patron, bloque, re.IGNORECASE)
    if match:
campos[campo] = match.group(1).strip()
    return campos

def seleccionar_bloque_valido(texto):
    bloques = extraer_bloques_encabezado(texto)
    for bloque in reversed(bloques):
        campos = extraer_campos_desde_bloque(bloque)
        if 'contabilidadfeci@elcorteingles.es' not in campos.get('De', '').lower():
            return campos
    return {'De': '', 'Enviado': '', 'Para': '', 'Cc': '', 'Asunto': ''}

def palabras_significativas(texto):
    return set([p.lower() for p in re.findall(r'\b\w{4,}\b', texto)])

def asignar_acreedor(dominio, df_acreedores, acreedores_no_asignados):
    fila = df_acreedores[df_acreedores["DOMINIO_ACREEDOR"] == dominio]
    if not fila.empty:
        return fila.iloc[0]["ACREEDOR"]
    for acreedor in acreedores_no_asignados:
        if palabras_significativas(dominio) & palabras_significativas(acreedor):
            acreedores_no_asignados.remove(acreedor)
            return acreedor
    mejor = process.extractOne(dominio, acreedores_no_asignados, scorer=fuzz.partial_ratio)
    if mejor and mejor[1] >= 50:
        acreedores_no_asignados.remove(mejor[0])
        return mejor[0]
    return ""

def manejar_casuistica(campos, nombre_archivo, contenido, permitir_guardado):
    def ejecutar_opcion(opcion):
        nonlocal campos, nombre_archivo, contenido
        campos['CASUISTICA'] = opcion
        campos['REVISADO_POR'] = usuario
campos['FECHA_REVISION'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if opcion == "Nueva Factura" and permitir_guardado:
        if messagebox.askyesno("Guardar adjunto", f"¿Deseas guardar el adjunto '{nombre_archivo}' en la carpeta del acreedor?"):
            carpeta_acreedor = os.path.join(carpeta_descarga, campos['ACREEDOR'] if campos['ACREEDOR'] else "desconocido")
            os.makedirs(carpeta_acreedor, exist_ok=True)
            destino = os.path.join(carpeta_acreedor, nombre_archivo)
            with open(destino, 'wb') as f:
                f.write(base64.b64decode(contenido))
    elif opcion == "Aprobación":
        messagebox.showinfo("En desarrollo", "⚙️ Funcionalidad aún en desarrollo.")
    elif opcion == "Otros":
        comentario = simpledialog.askstring("Otros", "Describe la casuística:")
        campos["OBSERVACION"] = comentario if comentario else ""
    elif opcion == "Modificar Acreedor":
        nuevo_acreedor = simpledialog.askstring("Modificar Acreedor", "Introduce el nuevo acreedor:")
        if nuevo_acreedor:
            campos['ACREEDOR'] = nuevo_acreedor
            text_campos.delete(1.0, tk.END)
            for k, v in campos.items():
                text_campos.insert(tk.END, f"{k}: {v}\n")
            mostrar_menu_casuistica()
            return
    avanzar_event.set(True)

def mostrar_menu_casuistica():
    top = tk.Toplevel()
    top.title("Selecciona casuística")
    tk.Label(top, text="¿Qué tipo de casuística aplica?").pack(pady=10)
    for texto in ["Nueva Factura", "Aprobación", "Otros", "Modificar Acreedor"]:
        tk.Button(top, text=texto, width=20, command=lambda t=texto: [top.destroy(), ejecutar_opcion(t)]).pack(pady=2)

    mostrar_menu_casuistica()
    root.wait_variable(avanzar_event)

def iniciar_revision():
    try:
        n = int(entry_n.get())
    except ValueError:
        messagebox.showerror("Error", "Introduce un número válido.")
        return

    text_campos.delete(1.0, tk.END)
    text_cuerpo.delete(1.0, tk.END)
app = PublicClientApplication(client_id=client_id, authority=f'https://login.microsoftonline.com/{tenant_id}')
    token_response = app.acquire_token_interactive(scopes=['Mail.Read'])
    if 'access_token' not in token_response:
        messagebox.showerror("Error", f"No se pudo obtener token: {token_response.get('error_description')}")
        return
    access_token = token_response['access_token']
    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}
endpoint = f'https://graph.microsoft.com/v1.0/me/mailFolders/{carpeta_id}/messages?$top={n}&$orderby=receivedDateTime desc'
    response = requests.get(endpoint, headers=headers)
    if response.status_code != 200:
        messagebox.showerror("Error", f"No se pudieron obtener correos: {response.status_code}")
        return
    messages = response.json().get('value', [])
    dataset = []
    acreedores_path = os.path.join("C:/FACTURAS", "ACREEDORES.xlsx")
    df_acreedores = pd.read_excel(acreedores_path, engine="openpyxl")
    df_acreedores.columns = ["ACREEDOR", "DOMINIO_ACREEDOR"]
    df_acreedores["DOMINIO_ACREEDOR"] = df_acreedores["DOMINIO_ACREEDOR"].astype(str).str.lower()
    acreedores_no_asignados = df_acreedores["ACREEDOR"].astype(str).tolist()

    for msg in messages:
        avanzar_event.set(False)
        message_id = msg['id']
        tiene_adjunto = 'SI' if msg.get('hasAttachments', False) else 'NO'
cuerpo_url = f"https://graph.microsoft.com/v1.0/me/messages/{message_id}?$select=body"
    cuerpo_resp = requests.get(cuerpo_url, headers=headers)
    if cuerpo_resp.status_code == 200:
        html_cuerpo = cuerpo_resp.json().get('body', {}).get('content', '')
        soup = BeautifulSoup(html_cuerpo, 'html.parser')
        cuerpo = soup.get_text(separator='\n')
    else:
        cuerpo = "[Error al obtener cuerpo]"
    campos = seleccionar_bloque_valido(cuerpo)
    campos['TIENE_ADJUNTO'] = tiene_adjunto
    dominio = extraer_dominio(campos['De']) if campos['De'] else 'desconocido'
    campos['DOMINIO'] = dominio
    campos['ACREEDOR'] = asignar_acreedor(dominio, df_acreedores, acreedores_no_asignados)
    campos['ID_CORREO'] = message_id

    text_campos.delete(1.0, tk.END)
    text_cuerpo.delete(1.0, tk.END)
    for k, v in campos.items():
        text_campos.insert(tk.END, f"{k}: {v}\n")
    text_cuerpo.insert(tk.END, cuerpo)

    contenido = None
    nombre_archivo = ""
    if tiene_adjunto == 'SI':
adjuntos_url = f"https://graph.microsoft.com/v1.0/me/messages/{message_id}/attachments"
    adjuntos_resp = requests.get(adjuntos_url, headers=headers)
    if adjuntos_resp.status_code == 200:
        for adjunto in adjuntos_resp.json().get('value', []):
            if '@odata.mediaContentType' in adjunto:
                nombre_archivo = adjunto.get('name', 'adjunto')
                if nombre_archivo.lower().endswith('.png'):
                    continue
                contenido = adjunto.get('contentBytes', '')
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                    tmp.write(base64.b64decode(contenido))
tmp_path = tmp.name
    subprocess.Popen(['start', tmp_path], shell=True)
    break

manejar_casuistica(campos, nombre_archivo, contenido, permitir_guardado=(tiene_adjunto == 'SI'))
dataset.append(campos)
    
        df = pd.DataFrame(dataset)
        for col in ['OBSERVACION', 'CASUISTICA', 'REVISADO_POR', 'FECHA_REVISION', 'ID_CORREO']:
            if col not in df.columns:
                df[col] = ''
        df.to_excel("correos_extraidos.xlsx", index=False)
        messagebox.showinfo("Finalizado", "✅ Revisión completada y Excel generado.")
    
btn_iniciar.config(command=iniciar_revision)
root.mainloop()