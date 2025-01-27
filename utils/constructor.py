from .data_processor import DataProcessor 
from .names_processor import NamesProcessor
from .excel_processor import ExcelProcessor


class Constructor:
    def __init__(self, file_path, output_path, data_path, row_processor):
        self.file_path = file_path
        self.output_path = output_path
        self.data_path = data_path
        self.row_processor = row_processor  # Instancia que contiene el método process_row
        
        # Instancias de procesadores
        self.id_processor = ExcelProcessor(self.file_path)
        self.name_processor = ExcelProcessor(self.data_path)
        self.names_processor = None
        self.id_names_processor = None
        self.data_processor = None  
    
    def initialize_processors(self):
        # Leer y cargar datos del archivo principal
        self.id_processor.read_excel()
        id_sheet = self.id_processor.get_sheet()
        self.names_processor = NamesProcessor(id_sheet)

        # Leer y cargar datos del archivo de nombres
        self.name_processor.read_excel()
        data_sheet = self.name_processor.get_sheet()
        self.id_names_processor = NamesProcessor(data_sheet)
        self.data_processor = DataProcessor(id_sheet, data_sheet)
    
    def iterate_columns(self):
        if not self.names_processor or not self.id_names_processor:
            raise ValueError("Procesadores no inicializados. Llama a 'initialize_processors' primero.")
        
        column_letters = ['H','I','J']  # Lista de columnas a procesar
        all_formats = []

        for col_letter in column_letters:
            # Obtener los valores de la columna correspondiente
            column_data2= self.names_processor.get_data_column(start_column=col_letter)
            names = self.names_processor.extract_names(column_data2)
            ids = self.id_names_processor.map_names_to_ids(names)
            
            column_data = self.data_processor.get_data_column(start_column=col_letter)
            print(f"Datos obtenidos de la columna {col_letter}: {column_data}: {column_data2}")
            
            # Delegar el procesamiento de filas al método process_row existente
            formats = self.row_processor.process_row(ids, column_data)
            all_formats.extend(formats)
            print(f"Datos procesados para la columna {col_letter}")
        
        return all_formats


