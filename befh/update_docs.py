import os
import sys
import json
import pygsheets

from pprint import pprint
from pygsheets.exceptions import RequestError


class ArbitrageDocs(object):
    def __init__(self):
        self.project_root = os.path.abspath(os.path.dirname(__file__))
        self.gc = None

    def update_doc(self, exchange, instmt, price):
        while True:
            count = 0
            try:
                exchange_cell = self.get_cell_number(exchange, instmt)

                if not self.gc:
                    self.auth_sheets()

                doc = self.gc.open_by_key("1N1MDj7mXLbv_-LUj1peug53P9a2gh5jdD5BnPrOQVLU")

                sheet = doc.worksheet_by_title("ArbitrageTable")

                new_cell = doc.Cell()
                sheet.update_cell(exchange_cell, price)
                return True

            except RequestError as e:
                print("Encountered Request error: %s" % e)
                count += 1

            if count > 10:
                return False

    def get_cell_number(self, exchange, instmt):
        cells_config = self.project_root + os.sep + "exchange_cells.json"
        with open(cells_config, 'r') as read:
            cells = json.load(read)
        pprint(cells)

        return cells[exchange][instmt]['price']

    def auth_sheets(self, path_to_creds=None):
        if not path_to_creds:
            if sys.platform == "win32":
                credentials = "C:\\Users\\Wilhelm\\PycharmProjects\\arbitrage\\config\\service_creds.json"
            else:
                credentials = "/home/guldenmw/PycharmProjects/arbitrage/config/service_creds.json"
        else:
            credentials = path_to_creds

        self.gc = pygsheets.authorize(service_file=credentials, no_cache=True)

    # def update_exchanges_sheet(self,):

