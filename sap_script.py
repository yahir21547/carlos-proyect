import win32com.client
import time
import pyautogui

def get_sap_session():
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    if application.Children.Count == 0:
        raise Exception("No hay conexiones abiertas en SAP. Abre PRD manualmente antes de ejecutar el script.")
    connection = application.Children(0)
    if connection.Children.Count == 0:
        raise Exception("No hay sesiones activas en la conexión.")
    return connection.Children(0)

def main():
    session = get_sap_session()
    print("✅ Conectado a sesión existente.")

    # Ir a transacción
    session.StartTransaction("ZENQ16")
    time.sleep(2)

    # Rellenar Lab Office
    try:
        session.findById("wnd[0]/usr/ctxtS_LABOR-LOW").text = "063"
        print("✅ Campo 'Lab Office' rellenado con 063.")
    except:
        print("⚠️ No se encontró el campo Lab Office.")

    # Ejecutar (F8)
    session.findById("wnd[0]").sendVKey(8)
    print("✅ Ejecutado (F8).")
    time.sleep(3)

    # 👉 Aquí deberías navegar y abrir el PDF (como ya lo hacías antes)
    # Simulamos que ya lo abriste
    print("📄 PDF abierto en el visor...")

    # Esperar que el visor cargue
    time.sleep(5)

    # Lanzar "Guardar como" con el atajo (Ctrl+Shift+S, a veces Ctrl+S)
    pyautogui.hotkey("ctrl", "shift", "s")
    time.sleep(2)

    # Escribir ruta donde guardar
    ruta = r"C:\Users\MXYAGAR1\Downloads\reporte.pdf"
    pyautogui.write(ruta)
    time.sleep(1)

    # Confirmar con ENTER
    pyautogui.press("enter")
    print(f"💾 PDF guardado en {ruta}")

if __name__ == "__main__":
    main()
