import os
import sys
import time
import re
import json
import pygsheets
import pygsheets.utils as sheet_utils

import logging
from pprint import pprint
from pygsheets.exceptions import RequestError

logging.basicConfig(level=logging.WARNING)
log = logging.getLogger(__name__)


class ArbitrageDoc(object):
    def __init__(self, sheet_key=None):
        self.project_root = os.path.abspath(os.path.dirname(__file__))
        self._service_creds = os.path.join(self.project_root, "service_creds.json")
        self.exchange_cells = os.path.join(self.project_root, "exchange_cells.json")

        self.sheet_key = sheet_key if sheet_key else "1N1MDj7mXLbv_-LUj1peug53P9a2gh5jdD5BnPrOQVLU"

        self.gc = pygsheets.authorize(service_file=self._service_creds, no_cache=True)
        self.doc = self.gc.open_by_key(self.sheet_key)
        self.sheet = None

    # def update_doc(self, exchange, instmt, price):
    #     while True:
    #         count = 0
    #         try:
    #             exchange_cell = self.get_cell_number(exchange, instmt)
    #
    #             self.sheet = self.doc.worksheet_by_title("ArbitrageTable")
    #
    #             self.sheet.update_cell(exchange_cell, price)
    #             return True
    #
    #         except RequestError as e:
    #             print("Encountered Request error: %s" % e)
    #             count += 1
    #
    #         if count > 10:
    #             return False

    def get_cell_number(self, exchange, instmt):
        cells_config = self.project_root + os.sep + "exchange_cells.json"
        with open(cells_config, 'r') as read:
            cells = json.load(read)
        pprint(cells)

        return cells[exchange][instmt]['price']

    # def parse_cell(self, cell):
    #     raw_regex = r'(?P<column>[A-Za-z]+)(?P<row>\d+)'
    #     result = re.search(raw_regex, cell)
    #     print(result)
    #     return result.groupdict()

    def update_exchange_tables(self, force=False):
        with open(self.exchange_cells, 'r+') as rw:
            sheet_info = json.load(rw)

        exchanges = []
        for exchange in sheet_info['exchanges']:
            for name in exchange:
                exchanges.append(name)

        start_cell_label = sheet_info['fields']['start']
        start_cell_tuple = sheet_utils.format_addr(start_cell_label)
        row_spread = sheet_info['fields']['row_spread']
        column_spread = sheet_info['fields']['column_spread']

        currencies = sheet_info['fields']['currencies']
        currencies_len = len(currencies)
        currency_info_rows = sheet_info['fields']['currency_info_rows']

        fiat = sheet_info['fields']['fiat']
        fiat_len = len(fiat)
        info_columns = sheet_info['fields']['info_columns']

        catagories_color = tuple(sheet_info['fields']['catagories_color'])
        currency_row_color = tuple(sheet_info['fields']['currency_rows_color'])
        currency_info_rows_color = tuple(sheet_info['fields']['currency_info_rows_color'])

        sheet = self.doc.worksheet_by_title("ExchangesInfo")
        sheet.adjust_column_width(start=start_cell_tuple[1]-1, end=start_cell_tuple[1], pixel_size=90)
        sheet.adjust_column_width(start=start_cell_tuple[1], end=start_cell_tuple[1]+26, pixel_size=75)

        current_exch_label = start_cell_label
        current_exch_tuple = sheet_utils.format_addr(current_exch_label)

        for exchange in exchanges:
            log.warning(
                "\ncurrent_exch_tuple: {}\ncurrent_exch_label: {}".format(current_exch_tuple, current_exch_label))

            exchange_update_cell_list = []
            exchange_info = sheet_info['exchanges'][exchanges.index(exchange)][exchange]

            exchange_cell = pygsheets.Cell(current_exch_label, val=exchange, worksheet=sheet)
            # exchange_update_cell_list.append(exchange_cell)
            exchange_cell.set_text_alignment(alignment="MIDDLE")
            exchange_cell.set_text_alignment(alignment="CENTER")

            table_start = current_exch_label
            left_bottom = current_exch_tuple[0] + currencies_len * 3
            right_top = current_exch_tuple[1] + 1 + fiat_len + len(info_columns)
            table_end = sheet_utils.format_addr((left_bottom, right_top))

            log.warning("\ntable_start: {}\nleft_bottom: {}\nright_top: {}\ntable_end: {}".format(
                table_start,
                left_bottom,
                right_top,
                table_end
            ))

            table_range = pygsheets.DataRange(start=table_start, end=table_end, worksheet=sheet)
            # table_range.applay_format(exchange_cell)
            exchange_cell.color = tuple(exchange_info['info']['cell_color'])

            table_range.update_borders()

            # exchange_cell.value = exchange

            current_col_tuple = (current_exch_tuple[0], current_exch_tuple[1] + 1)
            current_col_label = sheet_utils.format_addr(current_col_tuple)

            current_col_cell = pygsheets.Cell(current_col_label, val="Category", worksheet=sheet)

            current_col_cell.color = catagories_color
            current_col_cell.set_text_alignment(alignment="MIDDLE")
            current_col_cell.set_text_alignment(alignment="CENTER")

            column_range = pygsheets.DataRange(
                start=current_col_label,
                end=sheet_utils.format_addr((current_col_tuple[0], right_top)),
                worksheet=sheet
            )
            column_range.applay_format(current_col_cell)

            exchange_update_cell_list.append(current_col_cell)
            # current_col_cell.value = "Category"

            time.sleep(1)

            current_col_tuple = (current_col_tuple[0], current_col_tuple[1] + 1)
            # fiat_columns_cell_list = []
            # fiat_columns_cell_list.append(current_col_cell)
            for i in range(0, fiat_len+len(info_columns)):
                labels = fiat + info_columns
                current_tuple = (current_col_tuple[0], i+current_col_tuple[1])
                current_label = sheet_utils.format_addr(current_tuple)

                exchange_update_cell_list.append(pygsheets.Cell(current_label, val=labels[i], worksheet=sheet))

                # exchange_cell.value = labels[i]

            # sheet.update_cells(cell_list=fiat_columns_cell_list)

            current_row_tuple = (current_exch_tuple[0]+1, current_exch_tuple[1])
            current_row_label = sheet_utils.format_addr(current_row_tuple)

            current_cell = pygsheets.Cell(current_row_label, worksheet=sheet)
            current_cell.color = currency_row_color
            current_cell.set_text_alignment(alignment="MIDDLE")
            current_cell.set_text_alignment(alignment="CENTER")

            range_start = current_row_label
            range_end = sheet_utils.format_addr((current_row_tuple[0] + currencies_len * 3 - 1, current_row_tuple[1]))
            currencies_range = pygsheets.DataRange(start=range_start, end=range_end, worksheet=sheet)
            currencies_range.applay_format(current_cell)

            time.sleep(1)

            # currency_rows_cell_list = []
            for currency in currencies:
                current_cell = pygsheets.Cell(current_row_label, val=currency, worksheet=sheet)
                # current_cell.value = currency
                exchange_update_cell_list.append(current_cell)

                current_row_label = sheet_utils.format_addr((current_row_tuple[0] + 3, current_row_tuple[1]))
                current_row_tuple = sheet_utils.format_addr(current_row_label)

            # sheet.update_cells(cell_list=currency_rows_cell_list)

            current_row_tuple = (current_exch_tuple[0] + 1, current_exch_tuple[1])
            current_row_label = sheet_utils.format_addr(current_row_tuple)

            for i in range(0, currencies_len):
                currency_start = current_row_label
                currency_end = sheet_utils.format_addr((current_row_tuple[0] + 2, current_row_tuple[1]))
                currency_range = pygsheets.DataRange(start=currency_start, end=currency_end, worksheet=sheet)
                currency_range.merge_cells()

                current_row_label = sheet_utils.format_addr((current_row_tuple[0] + 3, current_row_tuple[1]))
                current_row_tuple = sheet_utils.format_addr(current_row_label)

            current_row_tuple = (current_exch_tuple[0] + 1, current_exch_tuple[1] + 1)
            current_row_label = sheet_utils.format_addr(current_row_tuple)

            current_cell = pygsheets.Cell(current_row_label, worksheet=sheet)
            current_cell.color = currency_info_rows_color
            current_cell.set_text_alignment(alignment="MIDDLE")
            current_cell.set_text_alignment(alignment="CENTER")

            range_start = current_row_label
            range_end = sheet_utils.format_addr((current_row_tuple[0] + currencies_len * 3 - 1, current_row_tuple[1]))
            currencies_range = pygsheets.DataRange(start=range_start, end=range_end, worksheet=sheet)
            currencies_range.applay_format(current_cell)

            time.sleep(1)

            # currency_info_rows_cell_list = []
            for i in range(0, currencies_len):

                for ii in range(0, len(currency_info_rows)):
                    current_cell_tuple = (current_row_tuple[0]+ii, current_row_tuple[1])
                    current_cell_label = sheet_utils.format_addr(current_cell_tuple)

                    current_cell = pygsheets.Cell(current_cell_label, val=currency_info_rows[ii], worksheet=sheet)
                    # current_cell.value = currency_info_rows[ii]
                    exchange_update_cell_list.append(current_cell)

                current_row_tuple = (current_row_tuple[0] + len(currency_info_rows), current_row_tuple[1])
                current_row_label = sheet_utils.format_addr(current_row_label)

            # sheet.update_cells(cell_list=currency_info_rows_cell_list)

            base_cell = pygsheets.Cell("A1", worksheet=sheet)
            base_cell.set_text_alignment(alignment="MIDDLE")
            base_cell.set_text_alignment(alignment="CENTER")
            # table_range.applay_format(base_cell)

            sheet.update_cells(cell_list=exchange_update_cell_list)

            time.sleep(10)

            current_exch_tuple = (current_exch_tuple[0] + 1 + currencies_len * 3 + row_spread, current_exch_tuple[1])
            current_exch_label = sheet_utils.format_addr(current_exch_tuple)

    def test_function(self):
        sheet = self.doc.worksheet_by_title("ExchangesInfo")
        exchange_cells = pygsheets.Cell("A1", worksheet=sheet)
        # exchange_cells = pygsheets.DataRange(start="A1", end="A3", worksheet=sheet)
        # exchange_cells.update_borders()
        pprint(exchange_cells.format)
        left_top = "C2"
        # currencies_len = 8
        # fiat_len = 9
        #
        # left_bottom = int(left_top[1:]) + currencies_len * 3
        #
        # right_top = int(left_top[1:]) + 1 + fiat_len + 3
        # sheet.adjust_column_width(start=0, end=26, pixel_size=75)
        # print(sheet_utils.format_addr(left_top))
        # print(right_top)

        # def update_exchanges_sheet(self,):

