import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
import win32com.client
import time

class PDFtoWordConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Conversor PDF a Word")
        self.root.geometry("600x300")
        
        # Variables para almacenar rutas
        self.input_pdf = tk.StringVar()
        self.output_docx = tk.StringVar()
        
        # Crear y configurar el marco principal
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Entrada para archivo PDF
        ttk.Label(main_frame, text="Archivo PDF:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.pdf_entry = ttk.Entry(main_frame, textvariable=self.input_pdf, width=50)
        self.pdf_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="Examinar", command=self.browse_pdf).grid(row=0, column=2, pady=5)
        
        # Entrada para archivo DOCX
        ttk.Label(main_frame, text="Guardar como DOCX:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.docx_entry = ttk.Entry(main_frame, textvariable=self.output_docx, width=50)
        self.docx_entry.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="Examinar", command=self.browse_docx).grid(row=1, column=2, pady=5)
        
        # Barra de progreso
        self.progress = ttk.Progressbar(main_frame, length=400, mode='determinate')
        self.progress.grid(row=2, column=0, columnspan=3, pady=20)
        
        # Botón convertir
        self.convert_btn = ttk.Button(main_frame, text="Convertir", command=self.convert_pdf_to_docx)
        self.convert_btn.grid(row=3, column=0, columnspan=3, pady=10)
        
        # Etiqueta de estado
        self.status_label = ttk.Label(main_frame, text="")
        self.status_label.grid(row=4, column=0, columnspan=3)

    def browse_pdf(self):
        filename = filedialog.askopenfilename(
            title="Seleccionar archivo PDF",
            filetypes=[("Archivos PDF", "*.pdf")]
        )
        if filename:
            self.input_pdf.set(filename)
            # Autosugerir nombre para archivo DOCX
            docx_name = os.path.splitext(filename)[0] + '.docx'
            self.output_docx.set(docx_name)

    def browse_docx(self):
        filename = filedialog.asksaveasfilename(
            title="Guardar como DOCX",
            defaultextension=".docx",
            filetypes=[("Documentos Word", "*.docx")]
        )
        if filename:
            self.output_docx.set(filename)

    def convert_pdf_to_docx(self):
        if not self.input_pdf.get() or not self.output_docx.get():
            messagebox.showerror("Error", "Por favor, seleccione los archivos de entrada y salida")
            return

        try:
            self.convert_btn.state(['disabled'])
            self.status_label.config(text="Iniciando conversión...")
            self.progress['value'] = 10
            self.root.update()

            # Obtener rutas absolutas
            pdf_path = os.path.abspath(self.input_pdf.get())
            docx_path = os.path.abspath(self.output_docx.get())

            self.status_label.config(text="Iniciando Microsoft Word...")
            self.progress['value'] = 20
            self.root.update()

            # Crear una instancia de Word
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False

            self.status_label.config(text="Abriendo documento PDF...")
            self.progress['value'] = 40
            self.root.update()

            # Abrir el PDF
            doc = word.Documents.Open(pdf_path)

            self.status_label.config(text="Guardando como DOCX...")
            self.progress['value'] = 70
            self.root.update()

            # Guardar como DOCX
            doc.SaveAs2(docx_path, FileFormat=16)  # 16 = formato DOCX

            # Cerrar el documento y Word
            doc.Close()
            word.Quit()
            
            self.progress['value'] = 100
            self.status_label.config(text=f"¡Conversión completada! ✅\nArchivo guardado: {self.output_docx.get()}")
            messagebox.showinfo("Éxito", "¡Conversión completada exitosamente!")
            
        except Exception as e:
            error_msg = str(e)
            if "Word.Application" in error_msg:
                error_msg = "Error: Microsoft Word no está instalado o no se puede acceder a él."
            messagebox.showerror("Error", 
                f"{error_msg}\n\n"
                "Sugerencias:\n"
                "1. Asegúrese de tener Microsoft Word instalado\n"
                "2. Cierre cualquier documento PDF abierto\n"
                "3. Verifique que el PDF no esté protegido\n"
                "4. Intente guardar en una ubicación diferente")
            self.status_label.config(text="Error en la conversión ❌")
        finally:
            self.convert_btn.state(['!disabled'])
            # Asegurar que Word se cierre en caso de error
            try:
                word.Quit()
            except:
                pass

def main():
    root = tk.Tk()
    app = PDFtoWordConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()


# © 2025 Pau aka Cyberp87
# Todos los derechos reservados.
# Licencia: Propiedad intelectual registrada (RPI España)
