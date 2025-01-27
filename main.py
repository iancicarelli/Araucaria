from utils.excel_processor import ExcelProcessor
from utils.data_processor import DataProcessor
from utils.names_processor import NamesProcessor

if __name__ == "__main__":
    file_path = r"C:\Users\HP\Desktop\formatoEntrada.xlsx"
    output_path = r"C:\Users\HP\Desktop\formatoSalida.xlsx"
    data_path = r"C:\Users\HP\Desktop\data.xlsx"
    """ TESTEO DE OBTENER NOMBRES Y DATA
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
    """
    file_processor = ExcelProcessor(file_path)
    file_processor.read_excel()
    file_shit = file_processor.get_sheet()
     # Leer el archivo de IDs y nombres (data.xlsx)
    id_processor = ExcelProcessor(data_path)
    id_processor.read_excel()
    id_sheet = id_processor.get_sheet()

    # Crear una instancia de DataProcessor
    data_processor = DataProcessor(file_shit, id_sheet)  # Usando el mismo id_sheet como ejemplo

    # Procesar las columnas
    # all_formats = data_processor.iterate_columns(id="some_id")  # Ajusta el ID según sea necesario
    test1 = data_processor.get_data_column()
        
    all_formats = data_processor.process_row('id',test1)

    """
    # Ver resultados
    print("Todos los formatos procesados:")
    for fmt in all_formats:
        print(fmt)
    all_formats =id_processor.write_excel(all_formats,output_path)
    """