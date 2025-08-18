# carlos-proyect

Aplicación sencilla para extraer datos de PDFs de motores y exportarlos a un archivo Excel.

## Uso
1. **Seleccionar Excel**: elige el archivo de destino donde se guardarán los datos.
2. **Configurar columnas**: asigna para cada campo la columna correspondiente. Esta configuración se guarda en `column_config.json` para reutilizarse.
3. **Seleccionar PDF**: carga un PDF y los datos extraídos se muestran en pantalla y se insertan en la primera fila disponible del Excel sin sobrescribir información existente.

Requiere las librerías `openpyxl` para manipular archivos `.xlsx` y
`pdfplumber` para extraer texto de los PDFs.
Ten en cuenta que `openpyxl` elimina las shapes o drawings al guardar, por lo que al abrir el archivo Excel mostrará **"Removed Part: /xl/drawings/drawing1.xml"**.
Si necesitas conservar dichas formas, considera alternativas como `xlwings` o `win32com`.
