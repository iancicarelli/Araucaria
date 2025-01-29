import unittest
from openpyxl import load_workbook
from utils.data_processor import DataProcessor
from utils.excel_processor import ExcelProcessor  # Importamos la clase para leer Excel

class TestDataProcessor(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        """Carga el archivo Excel una sola vez para todas las pruebas"""
        file_path = r"C:\Users\HP\Desktop\formatoEntrada.xlsx"
        data_path = r"C:\Users\HP\Desktop\DataFormato.xlsx"
        cls.excel_processor = ExcelProcessor(file_path)
        cls.excel_data = ExcelProcessor(data_path)
        cls.excel_data.read_excel()
        cls.data_sheet = cls.excel_data.get_sheet()

        cls.excel_processor.read_excel()
        cls.sheet = cls.excel_processor.get_sheet()

    def setUp(self):
        """ Configura el DataProcessor con la hoja de datos cargada """
        self.processor = DataProcessor(self.sheet,self.data_sheet )

    def test_process_uxe(self):
        """ Verifica que `process_uxe` retorne correctamente los códigos y UXE """
        result = self.processor.process_uxe()
        self.assertGreater(len(result), 0, "La lista de UXE no debe estar vacía")
        for code, uxe in result:
            self.assertIsInstance(code, (int,float), "El código debe ser un string")
            self.assertIsInstance(uxe, (int, float), "El UXE debe ser un número")

    def test_verify_amount(self):
        """ Verifica que `verify_amount` maneje correctamente datos numéricos y no numéricos """
        test_data = [10, "abc", 5.5, None, 0]
        expected = [10, 0, 5.5, 0, 0]  # Los no numéricos deben convertirse en 0
        result = self.processor.verify_amount(test_data)
        self.assertEqual(result, expected)

    def test_process_row(self):
        """ Prueba `process_row` asociando códigos con valores verificados """
        column_data = [20, 10, 30]  # Datos simulados
        processed_data = self.processor.process_row(999, column_data)
        
        # Verificamos que los datos procesados sean correctos
        self.assertGreater(len(processed_data), 0, "Debe haber datos procesados")

    def test_get_data_column(self):
        """ Verifica que `get_data_column` extraiga correctamente una columna desde Excel """
        result = self.processor.get_data_column(3, "H")
        self.assertGreater(len(result), 0, "La columna extraída no debe estar vacía")

if __name__ == "__main__":
    unittest.main()
