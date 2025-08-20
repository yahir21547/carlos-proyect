import win32com.client
import time
import re
import os
import shutil
import sys

def get_sap_session():
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    return connection.Children(0)

def mover_pdf(change_no, destino_dir):
    """Busca el PDF más reciente en Temp y lo copia a la carpeta destino."""
    temp_dir = r"C:\Users\MXYAGAR1\AppData\Local\Temp"
    destino = os.path.join(destino_dir, f"{change_no}.pdf")

    time.sleep(3)  # esperar a que SAP cree el archivo
    pdfs = [
        os.path.join(temp_dir, f) for f in os.listdir(temp_dir) if f.lower().endswith(".pdf")
    ]
    if not pdfs:
        print("⚠️ No se encontró ningún PDF en Temp.")
        return

    pdf_reciente = max(pdfs, key=os.path.getctime)
    shutil.copy(pdf_reciente, destino)
    print(f"💾 PDF guardado como: {destino}")

def buscar_label(session, change_no):
    """Busca en la tabla el label correspondiente a un Change No. específico."""
    usr = session.findById("wnd[0]/usr")
    for child in usr.Children:
        if child.Type == "GuiLabel" and child.Text.strip() == change_no:
            return child
    return None

def main(destino_dir):
    session = get_sap_session()
    print("✅ Conectado a sesión existente.")

    # Ir a transacción ZENQ16
    session.StartTransaction("ZENQ16")
    time.sleep(2)

    # Campo Lab Office = 063
    try:
        session.findById("wnd[0]/usr/ctxtS_LABOR-LOW").text = "063"
        print("✅ Campo 'Lab Office' rellenado con 063.")
    except:
        print("⚠️ No se encontró el campo 'Lab Office'.")

    # Ejecutar (F8)
    session.findById("wnd[0]").sendVKey(8)
    print("✅ Ejecutado (F8).")
    time.sleep(2)

    # Buscar todos los Change No. en la columna correcta
    usr = session.findById("wnd[0]/usr")
    change_numbers = []
    for child in usr.Children:
        if child.Type == "GuiLabel":
            text = child.Text.strip()
            match = re.search(r"lbl\[(\d+),(\d+)\]", child.Id)
            if match and text.isdigit() and len(text) >= 6:
                col = int(match.group(1))
                if col > 70:  # columna de Change No.
                    change_numbers.append(text)

    if not change_numbers:
        print("⚠️ No se encontraron Change Numbers en la lista.")
        return

    print(f"✅ Se encontraron {len(change_numbers)} Change No.: {change_numbers}")

    # Procesar uno por uno
    for change_no in change_numbers:
        try:
            print(f"➡️ Abriendo Change No. {change_no}")
            label = buscar_label(session, change_no)
            if not label:
                print(f"⚠️ No se encontró en pantalla el Change No. {change_no}")
                continue

            label.setFocus()
            session.findById("wnd[0]").sendVKey(2)  # doble click
            time.sleep(2)

            # Abrir PDF
            try:
                pdf_btn = session.findById("wnd[0]/usr/btnCUST_REQ_CONFIG_PDF")
                pdf_btn.press()
                print("✅ Botón PDF presionado (documento abierto).")
                mover_pdf(change_no, destino_dir)
            except:
                print(f"⚠️ No se pudo abrir PDF para {change_no}")

            # Cerrar visor PDF si está abierto
            try:
                session.findById("wnd[1]").close()
                print("⬅️ Cerrado visor PDF.")
            except:
                pass

            # Regresar dos veces
            session.findById("wnd[0]").sendVKey(15)  # F3
            time.sleep(1)
            session.findById("wnd[0]").sendVKey(15)  # otra vez
            time.sleep(2)
            print(f"⬅️ Cerrado Change No. {change_no}, regresando a lista.")

        except Exception as e:
            print(f"⚠️ Error procesando Change No. {change_no}: {e}")

if __name__ == "__main__":
    destino = sys.argv[1] if len(sys.argv) > 1 else r"C:\\Users\\MXYAGAR1\\Downloads"
    main(destino)
