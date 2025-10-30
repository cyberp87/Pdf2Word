# PDF to DOCX Converter (Interfaz GUI)

Convierte PDFs a DOCX manteniendo texto, formato e imágenes (según la versión usada).

## Descripción
Esta pequeña aplicación en Python incluye varias versiones de un conversor PDF→DOCX con interfaz gráfica (tkinter). Está pensada para uso en Windows. Hay varias implementaciones en el repositorio:

- `convertir_pdf_a_docx.py` — Versión inicial con `pdf2docx` y GUI.
- `convertir_pdf_a_docx_v2.py` — Versión alternativa usando `PyPDF2` + `python-docx` (más simple, extrae texto, útil si el PDF no necesita imágenes ni formato complejo).
- `convertir_pdf_a_docx_v3.py` — Intento con `aspose-words` (requiere atención a la licencia y API; puede mejorar la fidelidad).
- `convertir_pdf_a_docx_v4.py` — Versión optimizada con `pdf2docx` (intentos adicionales de manejo de errores).
- `convertir_pdf_a_docx_v5.py` — Versión que usa Microsoft Word (COM via `pywin32`) para la conversión. Recomendado para mantener formato e imágenes exactamente.

Lee `PREREQUISITES.md` antes de ejecutar para ver dependencias y requisitos del sistema.

## Características principales

- Interfaz gráfica simple para seleccionar el PDF de entrada y la ruta/nombre del DOCX de salida.
- Barra de progreso y mensajes de estado.
- Soporte (según versión) para:
  - Mantener formato y estilo
  - Mantener imágenes
  - Extraer solo texto (en la versión `v2`)
  - Usar Microsoft Word para conversión exacta (`v5`)

## Recomendación rápida

- Para fidelidad (imágenes + formato): usa `convertir_pdf_a_docx_v5.py` (necesitas Microsoft Word instalado y permisos para automatizarlo).
- Si quieres algo sin Word y con texto simple: `convertir_pdf_a_docx_v2.py`.
- Si `pdf2docx` falla con errores como `Rect object has no attribute get_area`, prueba `v2` o `v5`.

## Instalación (rápida)

1. Clona o sube este repositorio a tu máquina Windows.
2. Abre PowerShell en la carpeta del proyecto.
3. (Opcional) crea y activa un entorno virtual.

Instala dependencias (ver `requirements.txt`):

```powershell
# Usando el ejecutable Python detectado en tu sistema (ejemplo C:/Python313/python.exe)
python -m pip install -r requirements.txt
```

Si prefieres instalar paquetes manualmente (ejemplo):

```powershell
python -m pip install pdf2docx PyPDF2 python-docx aspose-words pywin32
```

> Nota: `aspose-words` puede requerir licencia para uso comercial o funcionalidades avanzadas.

## Uso

Ejecuta la versión que prefieras, por ejemplo (PowerShell):

```powershell
# Versión recomendada para fidelidad (requiere Word)
python convertir_pdf_a_docx_v5.py

# Versión que extrae texto (sin imágenes)
python convertir_pdf_a_docx_v2.py
```

La interfaz abrirá una ventana. Usa los botones "Examinar" para seleccionar el PDF y para elegir dónde guardar el DOCX. Presiona "Convertir" para iniciar.

## Ejemplos de flujo

1. `convertir_pdf_a_docx_v5.py` (Word): abre la interfaz → selecciona PDF → elige DOCX de salida → Convertir → Word hará la conversión y guardará el .docx.
2. `convertir_pdf_a_docx_v2.py` (PyPDF2): abre la interfaz → selecciona PDF con texto extraíble → Convertir → .docx con texto plano/estructurado.

## Troubleshooting (problemas comunes)

- Error: `'Rect' object has no attribute 'get_area'` — causados por ciertos tipos de objetos en PDFs que `pdf2docx` intenta leer. Soluciones:
  - Prueba `convertir_pdf_a_docx_v2.py` (extrae texto) o `convertir_pdf_a_docx_v5.py` (usa Word).
  - Asegúrate de que el PDF no esté protegido.
  - Si el PDF es una imagen escaneada, usa OCR primero (p. ej. Tesseract) para extraer texto.

- Error al ejecutar `v5` sobre COM/Word:
  - Asegúrate de tener Microsoft Word instalado.
  - Cierra instancias abiertas de Word antes de usar la aplicación.
  - Ejecuta PowerShell como usuario con permisos para automatizar Word.

## Archivos del repositorio

- `convertir_pdf_a_docx.py` — versión base
- `convertir_pdf_a_docx_v2.py` — PyPDF2 + python-docx (texto)
- `convertir_pdf_a_docx_v3.py` — intento con aspose-words
- `convertir_pdf_a_docx_v4.py` — optimizaciones con pdf2docx
- `convertir_pdf_a_docx_v5.py` — versión recomendada usando Microsoft Word (mantiene formato e imágenes)
- `README.md` — este archivo
- `PREREQUISITES.md` — lista de requisitos
- `requirements.txt` — manifiesto de dependencias

## Contribuir

1. Haz un fork del repositorio.
2. Crea una rama nueva para tu cambio.
3. Haz commits claros y descriptivos.
4. Abre un Pull Request explicando los cambios.

## Licencia

Añade aquí la licencia que prefieras (por ejemplo MIT). Si no deseas incluir ninguna, indícalo antes de publicar en GitHub.

---

Si quieres, puedo también crear un `README` en inglés o añadir badges (build, license). ¿Quieres que lo suba todo a GitHub y cree un repo remoto por ti?