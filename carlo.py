import tkinter as tk
from tkinter import filedialog, ttk
import pdfplumber
import re
from style_utils import ABB_COLORS, aplicar_colorimetria

def extraer_datos(pdf_path):
    datos = {
        "Catalog Number": None,
        "Power (HP)": None,
        "Speed (RPM)": None,
        "Phase": None,
        "Hertz": None,
        "Voltage": None,
        "Order Codes": []
    }

    with pdfplumber.open(pdf_path) as pdf:
        text = ""
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text.replace("\n", " ") + " "

    # --- Catalog Number ---
    match_catalog = re.search(r"Catalog\s+Number\s+([A-Z0-9\-]+)", text, re.IGNORECASE)
    if match_catalog:
        datos["Catalog Number"] = match_catalog.group(1)

    # --- POWER ---
    match_power = re.search(r"Power\s+([\d.]+)\s*HP", text, re.IGNORECASE)
    if not match_power:
        match_power = re.search(r"\b([\d.]+)\s*HP\b", text)
    if match_power:
        datos["Power (HP)"] = match_power.group(1)

    # --- SPEED ---
    match_speed = re.search(r"Speed\s*\(RPM\)\s*(\d+)", text, re.IGNORECASE)
    if not match_speed:
        match_speed = re.search(r"\b(\d{3,4})\b\s*(?:RPM|R\.P\.M\.?)", text, re.IGNORECASE)
    if not match_speed:
        match_speed = re.search(r"\b(900|1200|1500|1800|3000|3600)\b", text)
    if match_speed:
        datos["Speed (RPM)"] = match_speed.group(1)

    # --- PHASE / HERTZ / VOLTAGE ---
    phv_matches = re.findall(r"\b(\d*)\s*/\s*(\d*)\s*/\s*([\d/]+)\b", text)
    for phase_val, hertz_val, volt_val in phv_matches:
        if len(volt_val) == 4 and volt_val.startswith("20"):  # evitar fechas
            continue
        datos["Phase"] = phase_val if phase_val else None
        datos["Hertz"] = hertz_val if hertz_val else None
        datos["Voltage"] = volt_val if volt_val else None
        break

    # --- ORDER CODES ---
    order_codes = re.findall(r"NEMA Motor Modifications\s+([A-Z0-9]{2,3})\s*-", text)
    datos["Order Codes"] = list(dict.fromkeys(order_codes))  # quitar duplicados manteniendo orden

    return datos

def seleccionar_pdf():
    file_path = filedialog.askopenfilename(
        title="Selecciona un archivo PDF",
        filetypes=[("Archivos PDF", "*.pdf")]
    )
    if file_path:
        datos = extraer_datos(file_path)
        mostrar_datos(datos)
        print(datos)  # para depuración

def mostrar_datos(datos):
    texto_salida.delete(1.0, tk.END)
    texto_salida.insert(
        tk.END,
        (
            f"Catalog Number: {datos['Catalog Number']}\n"
            f"HP: {datos['Power (HP)']}\n"
            f"RPM: {datos['Speed (RPM)']}\n"
            f"Phase: {datos['Phase']}\n"
            f"Hz: {datos['Hertz']}\n"
            f"Volts: {datos['Voltage']}\n"
            f"Order Codes: {', '.join(datos['Order Codes'])}"
        ),
    )

# --- INTERFAZ ---
root = tk.Tk()
root.title("Extractor de Datos PDF")
aplicar_colorimetria(root)

btn_cargar = ttk.Button(root, text="Seleccionar PDF", command=seleccionar_pdf)
btn_cargar.pack(pady=10)

texto_salida = tk.Text(
    root,
    height=10,
    width=80,
    bg=ABB_COLORS["bg"],
    fg=ABB_COLORS["fg"],
)
texto_salida.pack(pady=10)

root.mainloop()

