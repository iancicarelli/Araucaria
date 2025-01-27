from utils.excel_processor import ExcelProcessor
from utils.data_processor import DataProcessor
from utils.names_processor import NamesProcessor

if __name__ == "__main__":
    file_path = r"C:\Users\HP\Desktop\formatoEntrada.xlsx"
    output_path = r"C:\Users\HP\Desktop\formatoSalida.xlsx"
    data_path = r"C:\Users\HP\Desktop\data.xlsx"

    # Leer el archivo de IDs y nombres (data.xlsx)
    id_processor = ExcelProcessor(file_path)
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

    #Encontrar coincidencias
    name_process = ExcelProcessor(data_path)
    name_process.read_excel()
    data_sheet = name_process.get_sheet()
    # Mapear nombres a IDs
    ids = []
    for name in names:
        id_names_processor = NamesProcessor(data_sheet)
        id = id_names_processor.map_names_to_ids(names)
        if id:
            ids.append(id)

    print("IDs mapeados:")
    print(ids)
