import os
import sys
import re
import json
import pygsheets

from pprint import pprint
from pygsheets.exceptions import RequestError


class ArbitrageDoc(object):
    def __init__(self, sheet_key=None):
        self.project_root = os.path.abspath(os.path.dirname(__file__))
        self._service_creds = os.path.join(self.project_root, "service_creds.json")
        self.exchange_cells = os.path.join(self.project_root, "exchange_cells.json")

        self.sheet_key = sheet_key if sheet_key else "1N1MDj7mXLbv_-LUj1peug53P9a2gh5jdD5BnPrOQVLU"

        self.gc = pygsheets.authorize(service_file=self._service_creds, no_cache=True)
        self.doc = self.gc.open_by_key(self.sheet_key)
        self.sheet = None

    def update_doc(self, exchange, instmt, price):
        while True:
            count = 0
            try:
                exchange_cell = self.get_cell_number(exchange, instmt)

                self.sheet = self.doc.worksheet_by_title("ArbitrageTable")

                self.sheet.update_cell(exchange_cell, price)
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

    def parse_cell(self, cell):
        raw_regex = r'(?P<column>[A-Za-z]+)(?P<row>\d+)'
        result = re.search(raw_regex, cell)
        print(result)
        return result.groupdict()

    def update_exchange_tables(self, force=False):
        with open(self.exchange_cells, 'r+') as rw:
            sheet_info = json.load(rw)

        exchanges = []
        for exchange in sheet_info['exchanges']:
            for name in exchange:
                exchanges.append(name)

        start_cell = sheet_info['fields']['start']
        row_spread = sheet_info['fields']['row_spread']
        column_spread = sheet_info['fields']['column_spread']

        currencies = sheet_info['fields']['currencies']
        currencies_len = len(currencies)
        currency_rows = sheet_info['fields']['currency_rows']

        fiat = sheet_info['fields']['fiat']
        fiat_len = len(fiat)
        info_columns = sheet_info['fields']['info_columns']

        catagories_color = tuple(sheet_info['fields']['catagories_color'])

        sheet = self.doc.worksheet_by_title("ExchangesInfo")

        current_exchange_cell = start_cell
        for exchange in exchanges:
            parse_current_cell = self.parse_cell(current_exchange_cell)
            exchange_info = sheet_info['exchanges'][exchanges.index(exchange)][exchange]
            # sheet.update_cell(current_exchange_cell, exchanges)
            exchange_cell = pygsheets.Cell(current_exchange_cell, worksheet=sheet)
            exchange_cell.value = exchange
            exchange_cell.color = tuple(exchange_info['info']['cell_color'])

            current_exchange_cell = "%s%d" % (parse_current_cell["column"],
                                              int(parse_current_cell['row']) + 1 + currencies_len * 3 + row_spread)


    # def update_exchanges_sheet(self,):

