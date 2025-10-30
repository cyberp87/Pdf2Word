# Prerrequisitos

Antes de ejecutar la aplicación, revisa los siguientes requisitos y recomendaciones.

## Sistema operativo

- Recomendado: Windows 10 o superior. Algunas versiones (v5) usan COM de Microsoft Word y solo funcionarán en Windows.

## Python

- Python 3.8+ (probado con Python 3.13). Asegúrate de tener `python` o `python3` en tu PATH.

## Paquetes Python

Instala las dependencias listadas en `requirements.txt`.

Recomendación (PowerShell):

```powershell
python -m pip install -r requirements.txt
```

Dependencias principales incluidas en `requirements.txt`:

- pdf2docx — conversión directa PDF→DOCX
- PyPDF2 — extracción de texto (v2)
- python-docx — para crear/salvar documentos .docx (v2)
- aspose-words — alternativa (ver nota de licencia)
- pywin32 — automatización de Microsoft Word (v5)

## Microsoft Word (opcional pero recomendado para fidelidad)

- Si quieres que la conversión mantenga exactamente imágenes y formato, instala Microsoft Word en la máquina donde ejecutes `convertir_pdf_a_docx_v5.py`.
- Word permite usar COM (a través de `pywin32`) para abrir un PDF y guardarlo como DOCX con alta fidelidad.

## Notas sobre `aspose-words`

- `aspose-words` es una librería potente para convertir entre formatos. Puede requerir licencia para uso comercial o para funciones avanzadas. Revisa su documentación y licencia antes de usarla en producción.

## Permisos y ejecución

- Ejecuta PowerShell con permisos de usuario normal en la mayoría de los casos. Para automatizar Word en entornos restringidos puede ser necesario ejecutar con permisos elevados o ajustar políticas de seguridad.
- Si `pywin32` no funciona correctamente, ejecuta el siguiente script de post-instalación (si es necesario):

```powershell
# Solo si pywin32 requiere registro de COM (rara vez necesario en instalaciones recientes)
python -m pip install pywin32
python -c "import pywin32_postinstall; pywin32_postinstall.install()"
```

## Problemas comunes y soluciones

- Error `pdf2docx` con `'Rect' object has no attribute 'get_area'`: usar `convertir_pdf_a_docx_v2.py` (texto) o `convertir_pdf_a_docx_v5.py` (Word) como alternativa.
- PDF escaneado (sin texto extraíble): realiza OCR primero (por ejemplo, con Tesseract) para crear un PDF con texto o extraer imagen+texto.
