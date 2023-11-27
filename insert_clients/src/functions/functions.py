from ..consts.password_d import password
from pyodbc import Connection, connect, Cursor
from openpyxl import open, Workbook
from openpyxl.worksheet.worksheet import Worksheet

def make_connection(server_name: str = "SERVER'S NAME") -> Connection:
    DRIVER_NAME = 'SQL SERVER'
    SERVER_NAME = server_name
    DATABASE_NAME = "DATABASE'S NAME"

    cnxn : Connection = connect(
        driver = DRIVER_NAME,
        server = SERVER_NAME,
        database = DATABASE_NAME,
        trust_connection = 'yes',
        # uid = ,
        pwd = password,
    )
    return cnxn

def open_excel(path:str, sheet_name:str) -> tuple[Workbook, Worksheet]:
    workbook: Workbook = open(path)
    worksheet: Worksheet = workbook.get_sheet_by_name(sheet_name)
    return (workbook, worksheet)

def create_worksheet(workbook: Workbook, worksheet_name: str) -> Worksheet | None:
    try:
        return workbook.create_sheet(worksheet_name)
    except:
        return None

def get_query(name:str, rnc:str, dir:str, sector:str, tele_propie:str) -> str:
    return f"""
        DECLARE @NO_CODCLI VARCHAR(5);

        SET @NO_CODCLI = (SELECT TOP 1 M.CODCLI
                        FROM MAECLIENTE M 
                        ORDER BY CAST(M.CODCLI AS INT)
                        DESC) + 1;

        INSERT INTO [dbo].[MAECLIENTE](
                [CODCLI],[CODZON],[NOMCLI],[RIFCLI],[DIRCLI],[CIUDAD],[ESTADO],[ESTATUS],[CIPROPIE],[TLFPROPIE],
                [SECTOR],[FECHAING],[CONDICION],[SALDO],[UFICLI],[CHPPAGOCLI],[CORTECLI],[CorteDia],[CONTRIBU],
                [CAJAPLAS],[LIMITESALDO],[CODCATEGORIA],[CODESTADO],[CODMUNICIPIO],[CODCIUDAD],[CODZONAMISCELANEOS],
                [OPERFAC])
            VALUES
                (@NO_CODCLI,'85','{name if name != '' and name != None else ''}','{rnc if rnc != '' and rnc != None else ''}','{dir if dir != '' and dir != None else ''}','',
                'Santo Domingo','Suspendido','','{tele_propie if tele_propie != '' and tele_propie != 'None' else ''}','{sector if sector != '' and sector != None else ''}',GETDATE(),'',0.00,0,0,0,GETDATE(),0,
                0,0.00,'999','00031','00162','00426','51','');
    """

def execute_query(query:str, connection: Connection) -> bool:
    try:
        cursor: Cursor = connection.execute(query)
        cursor.commit()
        return True
    except Exception as e:
        print(e)
        return False