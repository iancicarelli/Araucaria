from openpyxl.utils import column_index_from_string
from models.format import Format

class DataProcessor:
    def __init__(self, sheet,id_sheet):
        self.sheet = sheet
        self.id_sheet = id_sheet


    #ASOCIA EL VALOR TOMADO CON EL CODIGO 
    def process_row(self, id, column_data):
        try:
            formats = []
            # Obtener todos los códigos (columna B)
            codes = [
                row[1] for row in self.sheet.iter_rows(
                    min_row=3, max_row=self.sheet.max_row, values_only=True
                )
            ]
        
            # Verificar que la cantidad de códigos y datos coincidan
            if len(codes) != len(column_data):
                print("La cantidad de códigos no coincide con los datos obtenidos.")
                return []

            # Asociar cada código con su respectivo valor
            for code, value in zip(codes, column_data):
                if code is not None:  # Asegurarse de que el código no sea None
                    print(f"Asociando código {code} con valor {value}")
                    formats.append(self.create_format(code, value, id))

            return formats
        except Exception as e:
            print(f"Error al procesar las filas: {e}")
            return []

    def verify_amount(self, column_data):
        amount_found = []
        for cell in column_data:
            print(f"Valor de la celda: {cell}")  # Para diagnóstico
            # Verificar si el valor es un número (entero o decimal)
            if isinstance(cell, (int, float)):
                amount_found.append(cell)  # Si es numérico, agregarlo
            else:
                # Si no es un número, agregar 0
                amount_found.append(0)
            
        return amount_found



    def create_format(self, code, amount, name_id):
        # Crear un objeto Format con los valores recibidos
        print(f"Creando formato con: Código {code}, Monto {amount}, ID {name_id}")
        return Format(code, amount, name_id)   

    def iterate_columns(self, id):
        
        column_letters = ['H']  # Lista de columnas a procesar (puedes agregar más)
        all_formats = []

        for col_letter in column_letters:
            # Obtener los valores de la columna correspondiente
            column_data = self.get_data_column(row_number=2, start_column=col_letter)
            print(f"Datos obtenidos de la columna {col_letter}: {column_data}")
            # Llamar a process_row para cada columna
            formats = self.process_row(id, column_data)
            all_formats.extend(formats)
            print(f"Datos procesados para la columna {col_letter}: {formats}")
        
        return all_formats

    #BUSCA EN UNA FILA
    def get_data_column(self, row_number=3, start_column='I'):
        if not self.sheet:
            print("No se ha cargado ninguna hoja. Asegúrate de llamar a load_excel primero.")
            return []

        try:
            print(f"Intentando obtener valores desde la fila {row_number}, columna {start_column}")
        
            # Convertir la letra de la columna a índice (1 para 'A', 2 para 'B', etc.)
            start_col_idx = column_index_from_string(start_column)
        
            values = []
            for row in self.sheet.iter_rows(min_row=row_number, max_row=self.sheet.max_row, min_col=start_col_idx, max_col=start_col_idx):
             value = row[0].value  # Extraer el valor de la celda en esa columna
             print(f"Valor en columna ID {start_col_idx}: {value}")
             values.append(value)

            return values

        except Exception as e:
            print(f"Error al obtener valores de la fila {row_number} desde la columna {start_column}: {e}")
        return []

    
