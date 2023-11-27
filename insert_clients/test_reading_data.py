from pyodbc import Connection
from unittest import TestCase, main as maintest
from src.functions.functions import make_connection, open_excel, get_query, execute_query
from src.functions.functions import create_worksheet
from src.consts.consts import PATH_FILE, NAME_WORKSHEET, NEW_WORKSHEET
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

class TestReadingData(TestCase):
    name: str = 'Harry'
    rnc: str = 'RncHarry'
    dir: str = 'Dirección Harry'
    sector: str = 'Sector Harry'
    tele_propie: str = '8291234567'

    def test_connection(self):
        connection = make_connection()
        self.assertEqual(type(connection), Connection)

    def test_opening_the_file(self):
        workbook, worksheet = open_excel(PATH_FILE, NAME_WORKSHEET)
        self.assertEqual(type(workbook), Workbook)
        self.assertEqual(type(worksheet), Worksheet)

    def test_get_query(self):
        query = get_query(self.name, self.rnc, self.dir, self.sector, self.tele_propie)
        self.assertEqual(type(query), str)
        self.assertIn('Harry', query)
    
    def test_execute_query(self):
        connection = make_connection()
        query = get_query(self.name, self.rnc, self.dir, self.sector, self.tele_propie)
        result: bool = execute_query(query, connection)
        self.assertEqual(result, True)
    
    def test_create_worksheet(self):
        workbook, _ = open_excel(PATH_FILE, NAME_WORKSHEET)
        result = create_worksheet(workbook, NEW_WORKSHEET)
        if result != None:
            self.assertEqual(type(result), Worksheet)
            self.assertEqual(len(list(result.values)), 0)
        else:
            self.assertEqual(result, None)
    
    def test_append_new_rows_worksheet_new(self):
        workbook, _ = open_excel(PATH_FILE, NAME_WORKSHEET)
        result = create_worksheet(workbook, NEW_WORKSHEET)
        for d in ['d', 'd', 'd', 'dda', 'dd']:
            result.append([d])
        workbook.save(PATH_FILE)
    
if __name__ == '__main__':
    maintest()