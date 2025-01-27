from openpyxl import Workbook
from models.format import Format

class DataProcessor:
    def __init__(self, sheet,id_sheet):
        self.sheet = sheet
        self.id_sheet = id_sheet

    def process_row(self, id, column_data):
        try:
            formats = []
            # Iterar sobre las filas, comenzando desde la fila 3
            for row in self.sheet.iter_rows(min_row=3, max_row=self.sheet.max_row, values_only=True):
                code = row[1]  # Columna B (índice 1)
                amounts = self.verify_amount(column_data)  # Verificar los valores de la columna

                # Imprimir los valores de la columna B y los montos verificados
                print(f"Codigo: {code}, Montos: {amounts}")        
                
                # Si el código es válido y los montos no son cero
                if code and amounts:
                    for amount in amounts:
                        if amount != 0:  # Solo procesar montos distintos de 0
                            formats.append(self.create_format(code, amount, id))

            return formats
        except Exception as e:
            print(f"Error al procesar las filas: {e}")
            return []

    def verify_amount(self, column_data):
        
        amount_found = []
        for cell in column_data:
            print(f"Valor de la celda: {cell}")  # Para diagnóstico
            # Si el valor es None o no es un número, se considera como 0
            if cell is None or (isinstance(cell, str) and not cell.replace('.', '', 1).isdigit()):
                amount_found.append(0)
            else:
                try:
                    # Intentar convertir el valor a float
                    amount = float(cell)
                    amount_found.append(amount)  # Si es numérico, agregarlo
                except (ValueError, TypeError):
                    amount_found.append(0)  # Si no se puede convertir, agregar 0
                    
        return amount_found



    def create_format(self, code, amount, name_id):
        # Crear un objeto Format con los valores recibidos
        return Format(code, amount, name_id)   

    def iterate_columns(self, id):
        
        column_letters = ['H', 'I', 'J']  # Lista de columnas a procesar (puedes agregar más)
        all_formats = []

        for col_letter in column_letters:
            # Obtener los valores de la columna correspondiente
            column_data = self.get_data_column(row_number=2, start_column=col_letter)
            # Llamar a process_row para cada columna
            formats = self.process_row(id, column_data)
            all_formats.extend(formats)
            print(f"Datos procesados para la columna {col_letter}: {formats}")
        
        return all_formats


    
