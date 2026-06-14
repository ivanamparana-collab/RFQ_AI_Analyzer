"""
RFQ_AI_Analyzer
================

Este programa lee un archivo PDF de una solicitud de cotizacion (RFQ),
extrae su texto, lo analiza con OpenAI y exporta los resultados a Word y Excel.

Pensado para ejecutarse sin cambios en Google Colab o en un entorno local.
"""

# =============================================================================
# 1. IMPORTACION DE LIBRERIAS
# =============================================================================

import json
import os
from pathlib import Path
from typing import Any, Dict

import pandas as pd
from docx import Document
from openai import OpenAI
from pypdf import PdfReader


# =============================================================================
# 2. CONFIGURACION GENERAL
# =============================================================================

# Nombre del PDF que debe estar en la misma carpeta del script o subirse a Colab.
PDF_ENTRADA = "RFQ_001.pdf"

# Archivos de salida que generara el programa.
WORD_SALIDA = "analisis_rfq.docx"
EXCEL_SALIDA = "analisis_rfq.xlsx"

# Modelo solicitado para el analisis.
MODELO_OPENAI = "gpt-4o-mini"


# =============================================================================
# 3. EXTRACCION DE TEXTO DEL PDF
# =============================================================================

def extraer_texto_pdf(ruta_pdf: str) -> str:
    """
    Extrae el texto de todas las paginas de un archivo PDF usando pypdf.

    Args:
        ruta_pdf: Ruta del archivo PDF de entrada.

    Returns:
        Texto completo extraido del PDF.

    Raises:
        FileNotFoundError: Si el archivo PDF no existe.
        ValueError: Si el PDF no contiene texto legible.
    """
    ruta = Path(ruta_pdf)

    if not ruta.exists():
        raise FileNotFoundError(
            f"No se encontro el archivo {ruta_pdf}. "
            "En Google Colab, sube el PDF antes de ejecutar el script."
        )

    lector = PdfReader(str(ruta))
    textos_paginas = []

    for numero_pagina, pagina in enumerate(lector.pages, start=1):
        texto_pagina = pagina.extract_text() or ""
        texto_pagina = texto_pagina.strip()

        if texto_pagina:
            textos_paginas.append(f"--- Pagina {numero_pagina} ---\n{texto_pagina}")

    texto_completo = "\n\n".join(textos_paginas).strip()

    if not texto_completo:
        raise ValueError(
            "No se pudo extraer texto del PDF. "
            "Si el RFQ esta escaneado como imagen, primero debe aplicarse OCR."
        )

    return texto_completo


# =============================================================================
# 4. ANALISIS DEL RFQ CON OPENAI
# =============================================================================

def crear_prompt_analisis(texto_rfq: str) -> str:
    """
    Construye el prompt que se enviara al modelo de OpenAI.

    Args:
        texto_rfq: Texto extraido del PDF.

    Returns:
        Prompt completo con instrucciones y contenido del RFQ.