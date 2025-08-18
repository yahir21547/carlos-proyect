import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pdfplumber
import re
import json
from openpyxl import load_workbook
from style_utils import ABB_COLORS, aplicar_colorimetria

CONFIG_FILE = "column_config.json"
excel_path = None

DEFAULT_COLUMNS = {
    "Catalog Number": "A",
    "Power (HP)": "B",
    "Speed (RPM)": "C",
    "Phase": "D",
    "Hertz": "E",
    "Voltage": "F",
    "Order Codes": "G",
}


def load_column_config():
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            return {**DEFAULT_COLUMNS, **data}
    except Exception:
        return DEFAULT_COLUMNS.copy()


column_config = load_column_config()

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

def save_column_config():
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(column_config, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def configurar_columnas():
    config_win = tk.Toplevel(root)
    config_win.title("Configurar Columnas")
    aplicar_colorimetria(config_win)
    entries = {}
    for i, (campo, columna) in enumerate(column_config.items()):
        tk.Label(config_win, text=campo).grid(row=i, column=0, padx=5, pady=5, sticky="e")
        e = tk.Entry(config_win)
        e.insert(0, columna)
        e.grid(row=i, column=1, padx=5, pady=5)
        entries[campo] = e

    def guardar():
        for campo, entry in entries.items():
            column_config[campo] = entry.get().strip().upper() or DEFAULT_COLUMNS[campo]
        save_column_config()
        config_win.destroy()

    ttk.Button(config_win, text="Guardar", command=guardar).grid(
        row=len(column_config), columnspan=2, pady=10
    )


def seleccionar_excel():
    global excel_path
    excel_path = filedialog.askopenfilename(
        title="Selecciona un archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx")]
    )
    if excel_path:
        try:
            wb = load_workbook(excel_path)
            wb.close()
        except Exception as exc:
            messagebox.showerror("Error al abrir Excel", str(exc))
            excel_path = None


def _cell_is_empty(cell):
    """Return True if *cell* should be considered empty.

    Cells that are part of a formatted table can contain formulas even when no
    user data has been entered yet.  In such cases ``cell.value`` will hold the
    formula string (``'=...'``) and the previous implementation treated that as
    existing data.  By considering any formula-only cell as empty we correctly
    append new rows to tables without skipping thousands of rows.
    """

    return cell.value in (None, "") or cell.data_type == "f"


def find_next_empty_row(worksheet, columns):
    row = 2
    while True:
        if all(_cell_is_empty(worksheet[f"{col}{row}"]) for col in columns):
            return row
        row += 1


def guardar_en_excel(datos):
    if excel_path is None:
        messagebox.showwarning(
            "Excel no seleccionado", "Por favor selecciona un archivo Excel"
        )
        return

    wb = load_workbook(excel_path)
    ws = wb.active
    row = find_next_empty_row(ws, column_config.values())
    for campo, columna in column_config.items():
        valor = datos.get(campo)
        if campo == "Order Codes":
            valor = ", ".join(valor)
        ws[f"{columna}{row}"] = valor
    wb.save(excel_path)
    wb.close()


def seleccionar_pdf():
    file_path = filedialog.askopenfilename(
        title="Selecciona un archivo PDF",
        filetypes=[("Archivos PDF", "*.pdf")]
    )
    if file_path:
        datos = extraer_datos(file_path)
        mostrar_datos(datos)
        guardar_en_excel(datos)
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

btn_excel = ttk.Button(root, text="Seleccionar Excel", command=seleccionar_excel)
btn_excel.pack(pady=5)

btn_config = ttk.Button(root, text="Configurar Columnas", command=configurar_columnas)
btn_config.pack(pady=5)

btn_cargar = ttk.Button(root, text="Seleccionar PDF", command=seleccionar_pdf)
btn_cargar.pack(pady=5)

texto_salida = tk.Text(
    root,
    height=10,
    width=80,
    bg=ABB_COLORS["bg"],
    fg=ABB_COLORS["fg"],
)
texto_salida.pack(pady=10)

root.mainloop()

