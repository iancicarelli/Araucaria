import unittest
from openpyxl import load_workbook
from utils.constructor import Constructor
from utils.data_processor import DataProcessor
from utils.excel_processor import ExcelProcessor
from utils.names_processor import NamesProcessor

class TestConstructorWithFiles(unittest.TestCase):
    def setUp(self):
        # Rutas a los archivos Excel
        data_path = r"C:\Users\HP\Desktop\data.xlsx"
        file_path = r"C:\Users\HP\Desktop\formatoEntrada.xlsx"

        # Inicializar procesadores de Excel
        self.id_processor = ExcelProcessor(file_path)
        self.name_processor = ExcelProcessor(data_path)

        # Leer y cargar datos del archivo de IDs
        self.id_processor.read_excel()
        id_sheet = self.id_processor.get_sheet()

        # Inicializar NamesProcessor con la hoja de IDs
        self.names_processor = NamesProcessor(id_sheet)

        # Leer y cargar datos del archivo de nombres
        self.name_processor.read_excel()
        data_sheet = self.name_processor.get_sheet()

        # Inicializar NamesProcessor con la hoja de datos
        self.id_names_processor = NamesProcessor(data_sheet)

        # Inicializar DataProcessor con las hojas cargadas
        self.data_processor = DataProcessor(id_sheet, data_sheet)

        # Inicializar Constructor con los procesadores configurados
        self.constructor = Constructor(
            file_path=file_path,
            output_path="mock_output/",
            data_path=data_path,
            row_processor=self.data_processor,
        )

        # Inicializar procesadores del constructor
        self.constructor.initialize_processors()

    def test_get_last_column_with_value(self):
        # Probar que devuelve la última columna con valor
        result = self.constructor.get_last_column_with_value()
        print(f"Última columna con valor: {result}")
        self.assertIsNotNone(result, "Debe devolver una columna válida con valor.")
    
    def test_generate_column_letters_up_to(self):
        # Probar la generación de letras de columna
        last_column_letter = self.constructor.get_last_column_with_value()
        result = self.constructor.generate_column_letters_up_to(last_column_letter, start_column_letter="H")
        print(f"Letras de columna generadas: {result}")
        self.assertGreaterEqual(len(result), 1, "Debe generar al menos una columna.")
    
    def test_iterate_columns(self):
        # Ejecutar el método y verificar que no lance errores
        result = self.constructor.iterate_columns()
        #print(f"Resultado de iterate_columns: {result}")
        self.assertIsInstance(result, list, "El resultado debe ser una lista.")
        self.assertGreater(len(result), 0, "Debe haber al menos un resultado procesado.")
    

if __name__ == "__main__":
    unittest.main()