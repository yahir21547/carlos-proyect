# carlos-proyect

Aplicación sencilla para extraer datos de PDFs de motores y exportarlos a un archivo Excel.

## Uso
1. **Seleccionar Excel**: elige el archivo de destino donde se guardarán los datos. La ruta se guarda automáticamente y se reutiliza hasta que el usuario elija un nuevo archivo.
2. **Configurar columnas**: asigna para cada campo la columna correspondiente. Esta configuración se guarda en `column_config.json` para reutilizarse.
3. **Descargar y Actualizar**: con el botón *Descargar y Actualizar* se ejecuta el script `sap_script.py`, se procesan automáticamente todos los PDFs descargados y los datos se insertan en el Excel evitando duplicados.

Requiere las librerías `openpyxl` para manipular archivos `.xlsx`,
`pdfplumber` para extraer texto de los PDFs y `pyautogui` junto con
`pywin32` para interactuar con SAP.
Ten en cuenta que `openpyxl` elimina las shapes o drawings al guardar, por lo que al abrir el archivo Excel mostrará **"Removed Part: /xl/drawings/drawing1.xml"**.
Si necesitas conservar dichas formas, considera alternativas como `xlwings` o `win32com`.
