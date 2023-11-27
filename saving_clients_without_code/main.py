import pyodbc as pd
import pandas
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell
from pyodbc import Connection, Cursor
from password_d import password

def make_connection() -> Connection:
    DRIVER_NAME = 'SQL SERVER'
    SERVER_NAME = "SERVER'S NAME"
    DATABASE_NAME = "DATABASE'S NAME"
    
    connection_string = """
        DRIVER={{{DRIVER_NAME}}};
        SERVER={{{SERVER_NAME}}};
        DATABASE={{{DATABASE_NAME}}};
        Trust_Connection=yes;
        uid=sist-hisrael;
        pwd={{{password}}};
    """ 
    cnxn : Connection = pd.connect(
        driver = DRIVER_NAME,
        server = SERVER_NAME,
        database = DATABASE_NAME,
        trust_connection = 'yes',
        # uid = ,
        pwd = password,
    )
    return cnxn

def get_data(row_value, cnxn):
    query = f"""SELECT CODCLI, NOMCLI 
                    FROM MAECLIENTE
                    WHERE NOMCLI LIKE '%{row_value}%' AND CODEMPRESA = 'VY'"""
    return pandas.read_sql(query, cnxn)

def make_query(row_value: str, cnxn, times:int = 0):
    error: bool = False
    index: int = 0
    try:
        # ! Use to split the string.
        index = row_value.index(' ')
    except ValueError:
        error = True
    # ! Do the query.
    df_existing_tables = get_data(row_value, cnxn)
    if len(df_existing_tables.values) == 0:
        # ! Split and strip the data.
        new_row_value = row_value[index:].strip().removeprefix('- ').strip().replace('  ', ' ')

        # ! Do the query.
        df_existing_tables = get_data(new_row_value, cnxn)
        # ? If the query doesn't return data.
        if len(df_existing_tables.values) == 0:
            if error:
                return (new_row_value, None, times)
            return make_query(new_row_value, cnxn, times + 1)
        else:
            # ! When the query returns data.
            return (new_row_value, df_existing_tables.values, times + 1)
    else:
        # ! When the query returns data.
        return (row_value, df_existing_tables.values, times)

def main(working_book):
    values: list[tuple[str, list[list[str]]]] = []
    without_codcli = []
    books = []
    cnxn = make_connection()
    rutas = 'rutas.xlsx'
    workbook = openpyxl.load_workbook(rutas)
    
    for book in workbook:
        book: Worksheet = book
        books.append(book)
        if working_book == str(book):
            for row in book:
                row: Cell = row
                row_value: str = str(row[1].value).strip()
                get_province = 0
                
                provinces = [
                    'NAVARRETE',
                    'VILLA GONZALES',
                    'ESPERANZA',
                    'MAO',
                    'SANTIAGO RODRIGUEZ',
                    'VILLA LOS ALMACIDOS',
                    'DAJABÓN','Montecristi',
                    'TAMBORIL',
                    'LA ROMANA',
                    'HIGUEY''Bavaro',
                    'PUERTO PLATA',
                    'SOSUA CABARETE',
                    'LA ROMANA',
                    'HIGUEY',
                    'Bavaro'
                    'ENTREGADO POR:',
                    ]
                
                for province in provinces:
                    if province == row_value.strip():
                        get_province += 1
                        without_codcli.append(row_value)
                
                if get_province == 0:
                    (row_value2, values2, times) = make_query(row_value, cnxn)
                    if type(values2) == None:
                        without_codcli.append(row_value + " <-> " + row_value2)
                    elif row_value2 == '1' or row_value2 == '2':
                        without_codcli.append(row_value + " <-> " + row_value2)
                    else:
                        values.append((times, row_value + " <-> " + row_value2, values2))

    for book in books:
        print(book)
    for without in without_codcli:
        print(without)
    print("=============================================================================")
    # values.sort(key= lambda x: x[0])
    for value in values:
        print("=================================" + str(value[0]))
        print(value[1])
        print('')
        print(value[2])
    


NOEMI_URENA = '<Worksheet "Noemi Urena">'
ADRIANA = '<Worksheet "Adriana">'
ESTARLIN = '<Worksheet "Estarlin">'
BONIFACIO = '<Worksheet "Bonifacio">'
EDIXON_MATOS = '<Worksheet "Edixon Matos">'
JONATHAN_DE_LA_ROSA = '<Worksheet "Jonathan de la Rosa">'
MARIANO_LEBRON = '<Worksheet "Mariano Lebron">'
NATALI_CAPELLAN = '<Worksheet "Natali Capellan">'

main(BONIFACIO)