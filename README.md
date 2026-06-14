# RFQ_AI_Analyzer

Proyecto en Python para leer un archivo PDF de una solicitud de cotizacion (RFQ), extraer su contenido con `pypdf`, analizarlo con OpenAI usando `gpt-4o-mini` y exportar el resultado a Word y Excel.

## Funcionalidades

- Lee el archivo `RFQ_001.pdf`.
- Extrae texto del PDF usando `pypdf`.
- Envia el contenido a `gpt-4o-mini`.
- Identifica:
  - Alcance del proyecto.
  - Entregables.
  - Fechas importantes.
  - Riesgos.
  - Informacion faltante.
  - Supuestos.
  - Resumen ejecutivo.
- Exporta el analisis a `analisis_rfq.docx`.
- Exporta el analisis a `analisis_rfq.xlsx`.

## Estructura del proyecto

```text
RFQ_AI_Analyzer/
+-- main.py
+-- requirements.txt
+-- README.md
+-- RFQ_001.pdf
```

## Ejecucion en Google Colab

1. Sube los archivos `main.py`, `requirements.txt` y `RFQ_001.pdf` a tu entorno de Colab.

2. Instala las dependencias:

```python
!pip install -r requirements.txt
```

3. Configura tu API key de OpenAI:

```python
import os
os.environ["OPENAI_API_KEY"] = "tu_api_key"
```

4. Ejecuta el programa:

```python
!python main.py
```

5. Al finalizar se generaran:

```text
analisis_rfq.docx
analisis_rfq.xlsx
```

## Ejecucion local

1. Instala las dependencias:

```bash
pip install -r requirements.txt
```

2. Configura la variable de entorno `OPENAI_API_KEY`.

En PowerShell:

```powershell
$env:OPENAI_API_KEY="tu_api_key"
```

En macOS o Linux:

```bash
export OPENAI_API_KEY="tu_api_key"
```

3. Coloca `RFQ_001.pdf` en la carpeta del proyecto.

4. Ejecuta:

```bash
python main.py
```

## Nota sobre PDFs escaneados

Este proyecto usa `pypdf`, que extrae texto embebido en el PDF. Si el RFQ esta escaneado como imagen, sera necesario aplicar OCR antes de ejecutar el analisis.
