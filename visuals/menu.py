import tkinter as tk
from tkinter import filedialog, messagebox
from utils.constructor import Constructor
from utils.excel_processor import ExcelProcessor
from utils.data_processor import DataProcessor
from ttkthemes import ThemedTk
from tkinter import ttk,tk


class Menu:
    def __init__(self):
        # Ventana principal con tema "Arc"
        self.root = ThemedTk(theme="arc")
        self.root.title("Pino")
        self.root.geometry("500x500")
        
        # Cambiar el color de fondo de la barra superior
        self.root.configure(bg="#00953A")

        # Crear componentes de la interfaz
        self.create_widgets()
        
    def create_widgets(self):
        # Estilo general para los widgets
        style = ttk.Style()
        style.configure("TLabel", font=("Arial", 12), background="#00953A", foreground="white")
        style.configure("TButton", font=("Arial", 10))

        # Selección del archivo de entrada
        ttk.Label(self.root, text="Selecciona el archivo de entrada:").pack(pady=10)
        self.file_entry = ttk.Entry(self.root, width=40)
        self.file_entry.pack(pady=5)
        ttk.Button(self.root, text="Buscar", command=self.select_file).pack(pady=5)

        # Selección del archivo de datos
        ttk.Label(self.root, text="Selecciona el archivo de datos:").pack(pady=10)
        self.data_path_entry = ttk.Entry(self.root, width=40)
        self.data_path_entry.pack(pady=5)
        ttk.Button(self.root, text="Buscar", command=self.select_data_file).pack(pady=5)
        
        # Selección del directorio de salida
        ttk.Label(self.root, text="Selecciona el directorio de salida:").pack(pady=10)
        self.output_dir_entry = ttk.Entry(self.root, width=40)
        self.output_dir_entry.pack(pady=5)
        ttk.Button(self.root, text="Seleccionar", command=self.select_directory).pack(pady=5)

        # Nombre del archivo de salida
        ttk.Label(self.root, text="Nombre del archivo de salida:").pack(pady=10)
        self.output_name_entry = ttk.Entry(self.root, width=40)
        self.output_name_entry.pack(pady=5)
        
        # Botón para ejecutar
        ttk.Button(self.root, text="Ejecutar", command=self.run).pack(pady=20)
        
    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)

    def select_data_file(self):
        data_path = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        if data_path:
            self.data_path_entry.delete(0, tk.END)
            self.data_path_entry.insert(0, data_path)
        
    def select_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.output_dir_entry.delete(0, tk.END)
            self.output_dir_entry.insert(0, directory)
        
    def run(self):
        file_path = self.file_entry.get()
        data_path = self.data_path_entry.get()
        output_dir = self.output_dir_entry.get()
        output_name = self.output_name_entry.get()
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