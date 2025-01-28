from utils.excel_processor import ExcelProcessor
from utils.data_processor import DataProcessor
from utils.names_processor import NamesProcessor
from utils.constructor import Constructor  


def main():
    # Configuración de rutas
    file_path = r"C:\Users\HP\Desktop\formatoEntrada.xlsx"
    output_path = r"C:\Users\HP\Desktop\formatoSalida.xlsx"
    data_path = r"C:\Users\HP\Desktop\data.xlsx"

    # Crear una instancia de ExcelProcessor para obtener las hojas
    id_processor = ExcelProcessor(file_path)
    id_processor.read_excel()
    id_sheet = id_processor.get_sheet()  # Obtener la hoja de trabajo de id_processor

    # Crear una instancia de ExcelProcessor para el archivo de datos
    name_processor = ExcelProcessor(data_path)
    name_processor.read_excel()
    data_sheet = name_processor.get_sheet()  # Obtener la hoja de trabajo de name_processor

    # Crear una instancia de DataProcessor con las hojas obtenidas
    row_processor = DataProcessor(id_sheet, data_sheet)

    # Crear una instancia de Constructor
    constructor = Constructor(file_path, output_path, data_path, row_processor)

    # Inicializar procesadores
    constructor.initialize_processors()

    # Procesar columnas y obtener el resultado
    result = constructor.iterate_columns()
    id_processor.write_excel(result,output_path)  
  
if __name__ == "__main__":
    main()