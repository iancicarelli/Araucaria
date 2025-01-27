from utils.excel_processor import ExcelProcessor
from utils.data_processor import DataProcessor
from utils.names_processor import NamesProcessor

if __name__ == "__main__":
    file_path = r"C:\Users\HP\Desktop\formatoEntrada.xlsx"
    output_path = r"C:\Users\HP\Desktop\formatoSalida.xlsx"
    data_path = r"C:\Users\HP\Desktop\data.xlsx"
    
   # Leer el archivo principal (formatoEntrada.xlsx)
    excel_processor = ExcelProcessor(file_path)
    excel_processor.read_excel()
    sheet = excel_processor.get_sheet()

    # Leer el archivo de IDs y nombres (data.xlsx)
    id_processor = ExcelProcessor(data_path)
    id_processor.read_excel()
    id_sheet = id_processor.get_sheet()

    # Crear una instancia de NamesProcessor y cargar el archivo Excel
    processor = NamesProcessor(id_sheet)


     # Obtener datos de la columna
    data = processor.get_data_column(row_number=2, start_column='H')
    print("Valores obtenidos de la columna:")
    print(data)

    # Extraer nombres
    names = processor.extract_names(data)
    print("Nombres extraídos:")
    print(names)
    
    # Mapear nombres a IDs
    id_processor = ExcelProcessor(data_path)
    id_processor.read_excel()
    id_sheet = id_processor.get_sheet()
    
    id_names_processor = NamesProcessor(id_sheet)
    ids = [id_names_processor.map_names_to_ids([name]) for name in names]
    print("IDs mapeados:")
    print(ids)