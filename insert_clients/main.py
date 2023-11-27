from src.functions.functions import make_connection, open_excel, get_query, execute_query, create_worksheet
from src.consts.consts import PATH_FILE, NAME_WORKSHEET, NEW_WORKSHEET
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from pyodbc import Connection

def main():
    error = False
    workbook: Workbook = None
    worksheet: Worksheet = None
    worksheet_new: Worksheet = None
    try:
        # ! Obtener el workbook y el worksheet.
        workbook, worksheet = open_excel(PATH_FILE, NAME_WORKSHEET)
        worksheet_new = create_worksheet(workbook, NEW_WORKSHEET)
    except:
        error = True
    if not error:
        # ! Conectar con la base de datos.
        connection: Connection = make_connection(server_name="SERVER'S NAME")
        rows = list(worksheet.rows)
        for row in rows:
            name = row[0].value
            if name != 'Nombre':
                rnc = row[1].value
                dir = row[2].value
                # provincia = row[3].value
                # municipio = row[4].value
                sector = row[5].value
                # email = row[7].value
                # nom_propie = row[8].value
                # cedula_propie = row[9].value
                tele_propie = str(row[10].value)
                # * Obtener el query.
                if tele_propie != None:
                    tele_propie = tele_propie.replace('(', '').replace(')', '').strip()
                query = get_query(name=name, rnc=rnc, dir=dir, sector=sector, tele_propie=tele_propie)
                # ! Ejecutar el query.
                result = execute_query(query, connection)
                if not result:
                    worksheet_new.append([name, rnc, dir, sector, tele_propie])
                

        # ! Cerrar la conexión.
        connection.close()
        # ! Guardar el workbook.
        workbook.save(PATH_FILE)

if __name__ == '__main__':
    main()