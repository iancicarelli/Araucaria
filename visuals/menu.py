import tkinter as tk
from tkinter import filedialog, messagebox
from ttkthemes import ThemedTk
from tkinter import ttk
from utils.constructor import Constructor
from utils.excel_processor import ExcelProcessor
from utils.data_processor import DataProcessor


class Menu:
    def __init__(self):
        # Ventana principal con tema "Arc"
        self.root = ThemedTk(theme="arc")
        self.root.title("Pino")
        self.root.geometry("600x400")
        
        # Crear los widgets principales
        self.create_widgets()
        
    def create_widgets(self):
        # Crear un Notebook para pestañas
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Crear una pestaña principal para las entradas
        self.input_frame = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(self.input_frame, text="Configuración")

        # Variables para las rutas
        self.file_path_var = tk.StringVar()
        self.data_path_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        self.output_name_var = tk.StringVar()

        # Etiqueta principal
        ttk.Label(
            self.input_frame, text="Selección de directorios", font=("Arial", 16)
        ).grid(row=0, column=0, columnspan=3, pady=10)

        # Selección del archivo de entrada
        ttk.Label(self.input_frame, text="Archivo de entrada:").grid(
            row=1, column=0, sticky="e", pady=5
        )
        ttk.Entry(self.input_frame, textvariable=self.file_path_var, width=40).grid(
            row=1, column=1, padx=5, pady=5
        )
        ttk.Button(
            self.input_frame, text="Buscar", command=self.select_file
        ).grid(row=1, column=2, padx=5, pady=5)

        # Selección del archivo de datos
        ttk.Label(self.input_frame, text="Archivo de datos:").grid(
            row=2, column=0, sticky="e", pady=5
        )
        ttk.Entry(self.input_frame, textvariable=self.data_path_var, width=40).grid(
            row=2, column=1, padx=5, pady=5
        )
        ttk.Button(
            self.input_frame, text="Buscar", command=self.select_data_file
        ).grid(row=2, column=2, padx=5, pady=5)

        # Selección del directorio de salida
        ttk.Label(self.input_frame, text="Directorio de salida:").grid(
            row=3, column=0, sticky="e", pady=5
        )
        ttk.Entry(self.input_frame, textvariable=self.output_dir_var, width=40).grid(
            row=3, column=1, padx=5, pady=5
        )
        ttk.Button(
            self.input_frame, text="Seleccionar", command=self.select_directory
        ).grid(row=3, column=2, padx=5, pady=5)

        # Nombre del archivo de salida
        ttk.Label(self.input_frame, text="Nombre del archivo de salida:").grid(
            row=4, column=0, sticky="e", pady=5
        )
        ttk.Entry(self.input_frame, textvariable=self.output_name_var, width=40).grid(
            row=4, column=1, padx=5, pady=5
        )

        # Botón para ejecutar
        ttk.Button(
            self.input_frame, text="Ejecutar", command=self.run
        ).grid(row=5, column=0, columnspan=3, pady=20)

    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        if file_path:
            self.file_path_var.set(file_path)

    def select_data_file(self):
        data_path = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        if data_path:
            self.data_path_var.set(data_path)

    def select_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.output_dir_var.set(directory)

    def run(self):
        file_path = self.file_path_var.get()
        data_path = self.data_path_var.get()
        output_dir = self.output_dir_var.get()
        output_name = self.output_name_var.get()
        output_path = f"{output_dir}/{output_name}.xlsx"

        if not file_path or not data_path or not output_dir or not output_name:
            messagebox.showwarning("Advertencia", "Por favor, completa todos los campos.")
            return

        try:
            # Crear una instancia de ExcelProcessor para obtener las hojas
            id_processor = ExcelProcessor(file_path)
            id_processor.read_excel()
            id_sheet = id_processor.get_sheet()

            # Crear una instancia de ExcelProcessor para el archivo de datos
            name_processor = ExcelProcessor(data_path)
            name_processor.read_excel()
            data_sheet = name_processor.get_sheet()

            # Crear una instancia de DataProcessor con las hojas obtenidas
            row_processor = DataProcessor(id_sheet, data_sheet)

            # Crear una instancia de Constructor
            constructor = Constructor(file_path, output_path, data_path, row_processor)

            # Inicializar procesadores
            constructor.initialize_processors()

            # Procesar columnas y obtener el resultado
            result = constructor.iterate_columns()
            id_processor.write_excel(result, output_path)

            messagebox.showinfo("Éxito", f"Archivo generado en: {output_path}")
            self.root.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error: {e}")

    def show(self):
        self.root.mainloop()
        
        
  
