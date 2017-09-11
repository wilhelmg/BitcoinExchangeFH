import os
import sys
import time
import re
import json
import pygsheets
import pygsheets.utils as sheet_utils

import logging
from pprint import pprint, pformat
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
        self.sheet = self.doc.worksheet_by_title("ExchangesInfo")

        with open(self.exchange_cells, 'r+') as rw:
            self.sheet_info = json.load(rw)

        self.exchanges = []
        for exchange in self.sheet_info['exchanges']:
            for name in exchange:
                self.exchanges.append(name)

        self.start_cell_label = self.sheet_info['fields']['start']
        self.start_cell_tuple = sheet_utils.format_addr(self.start_cell_label)
        self.row_spread = self.sheet_info['fields']['row_spread']
        self.column_spread = self.sheet_info['fields']['column_spread']

        self.currencies = self.sheet_info['fields']['currencies']
        self.currencies_len = len(self.currencies)
        self.currency_info_rows = self.sheet_info['fields']['currency_info_rows']

        self.fiat = self.sheet_info['fields']['fiat']
        self.fiat_len = len(self.fiat)
        self.info_columns = self.sheet_info['fields']['info_columns']

        self.categories_color = tuple(self.sheet_info['fields']['catagories_color'])
        self.currency_row_color = tuple(self.sheet_info['fields']['currency_rows_color'])
        self.currency_info_rows_color = tuple(self.sheet_info['fields']['currency_info_rows_color'])

        self.current_exch_label = self.start_cell_label
        self.current_exch_tuple = sheet_utils.format_addr(self.current_exch_label)

        self.table_range = None

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

    # def validate_exch_tables(self):
    def find_exchange_label(self, exchange):
        exchange_start_cell = self.sheet.find(query=exchange)
        if not exchange_start_cell or len(exchange_start_cell) > 2:
            return False

        return exchange_start_cell[0].label

    def get_table_labels(self, exchange):
        tuples = self.get_table_tuples(exchange)
        if not tuples:
            return False

        return sheet_utils.format_addr(tuples[0]), \
               sheet_utils.format_addr(tuples[1]), \
               sheet_utils.format_addr(tuples[2]), \
               sheet_utils.format_addr(tuples[3])

    def get_table_tuples(self, exchange):
        start_label = self.find_exchange_label(exchange)
        if not start_label:
            return False

        lt = self.current_exch_tuple
        lb = (lt[0] + self.currencies_len * 3, lt[1])
        rt = (lt[0], lt[1] + 1 + self.fiat_len + len(self.info_columns))
        rb = (lb[0], rt[1])

        return lt, lb, rt, rb

    def validate_exch_tables(self):
        for exchange in self.exchanges:
            status = self.validate_exch_table(exchange)
            if status == "Missing":
                log.warning("%s table is missing. Creating table." % exchange)
                self.create_exchange_tables()
            elif status == "Invalid":
                log.warning("%s table is invalid. Recreating table." % exchange)
                self.create_exchange_table(exchange)
            log.debug("Successfully validated %s table." % exchange)

    def validate_exch_table(self, exchange):
        table_tuples = self.get_table_tuples(exchange)
        if not table_tuples:
            return "Missing"
        lt, lb, rt, rb = table_tuples[0], table_tuples[1], table_tuples[2], table_tuples[3]
        columns = self.sheet.get_values(start=lt, end=rt, include_empty=False)[0]
        raw_rows = self.sheet.get_values(start=lt, end=lb, include_empty=False)
        rows = [i[0] for i in raw_rows if i]
        missing_fiat = [i for i in self.fiat if i not in columns]
        missing_info = [i for i in self.info_columns if i not in columns]
        missing_currencies = [i for i in self.currencies if i not in rows]
        if missing_fiat or missing_info or missing_currencies:
            log.warning("\nMissing fiat columns: \n{}\nMissing Info Columns: \n{}\nMissing Currency Rows: \n{}".format(
                pformat(missing_fiat),
                pformat(missing_info),
                pformat(missing_currencies)

            ))
            return "Invalid"
        return "Valid"

    def create_exchange_table(self, exchange):
        exchange_update_cell_list = []
        exchange_info = self.sheet_info['exchanges'][self.exchanges.index(exchange)][exchange]

        exchange_cell = pygsheets.Cell(self.current_exch_label, val=exchange, worksheet=self.sheet)
        # exchange_update_cell_list.append(exchange_cell)
        exchange_cell.set_text_alignment(alignment="MIDDLE")
        exchange_cell.set_text_alignment(alignment="CENTER")

        table_start = self.current_exch_label
        left_bottom = self.current_exch_tuple[0] + self.currencies_len * 3
        right_top = self.current_exch_tuple[1] + 1 + self.fiat_len + len(self.info_columns)
        table_end = sheet_utils.format_addr((left_bottom, right_top))

        # log.warning("\ntable_start: {}\nleft_bottom: {}\nright_top: {}\ntable_end: {}".format(
        #     table_start,
        #     left_bottom,
        #     right_top,
        #     table_end
        # ))

        table_range = pygsheets.DataRange(start=table_start, end=table_end, worksheet=self.sheet)
        # table_range.applay_format(exchange_cell)
        exchange_cell.color = tuple(exchange_info['info']['cell_color'])

        table_range.update_borders()

        # exchange_cell.value = exchange

        current_col_tuple = (self.current_exch_tuple[0], self.current_exch_tuple[1] + 1)
        current_col_label = sheet_utils.format_addr(current_col_tuple)

        current_col_cell = pygsheets.Cell(current_col_label, val="Category", worksheet=self.sheet)

        current_col_cell.color = self.categories_color
        current_col_cell.set_text_alignment(alignment="MIDDLE")
        current_col_cell.set_text_alignment(alignment="CENTER")

        column_range = pygsheets.DataRange(
            start=current_col_label,
            end=sheet_utils.format_addr((current_col_tuple[0], right_top)),
            worksheet=self.sheet
        )
        column_range.applay_format(current_col_cell)

        # exchange_update_cell_list.append(current_col_cell)
        current_col_cell.value = "Category"

        time.sleep(1)

        current_col_tuple = (current_col_tuple[0], current_col_tuple[1] + 1)
        # fiat_columns_cell_list = []
        # fiat_columns_cell_list.append(current_col_cell)
        for i in range(0, self.fiat_len + len(self.info_columns)):
            labels = self.fiat + self.info_columns
            current_tuple = (current_col_tuple[0], i + current_col_tuple[1])
            current_label = sheet_utils.format_addr(current_tuple)

            exchange_update_cell_list.append(pygsheets.Cell(current_label, val=labels[i], worksheet=self.sheet))

            # exchange_cell.value = labels[i]

        # sheet.update_cells(cell_list=fiat_columns_cell_list)

        current_row_tuple = (self.current_exch_tuple[0] + 1, self.current_exch_tuple[1])
        current_row_label = sheet_utils.format_addr(current_row_tuple)

        current_cell = pygsheets.Cell(current_row_label, worksheet=self.sheet)
        current_cell.color = self.currency_row_color
        current_cell.set_text_alignment(alignment="MIDDLE")
        current_cell.set_text_alignment(alignment="CENTER")

        range_start = current_row_label
        range_end = sheet_utils.format_addr((current_row_tuple[0] + self.currencies_len * 3 - 1, current_row_tuple[1]))
        currencies_range = pygsheets.DataRange(start=range_start, end=range_end, worksheet=self.sheet)
        currencies_range.applay_format(current_cell)

        time.sleep(1)

        # currency_rows_cell_list = []
        for currency in self.currencies:
            current_cell = pygsheets.Cell(current_row_label, val=currency, worksheet=self.sheet)
            # current_cell.value = currency
            exchange_update_cell_list.append(current_cell)

            current_row_label = sheet_utils.format_addr((current_row_tuple[0] + 3, current_row_tuple[1]))
            current_row_tuple = sheet_utils.format_addr(current_row_label)

        # sheet.update_cells(cell_list=currency_rows_cell_list)

        current_row_tuple = (self.current_exch_tuple[0] + 1, self.current_exch_tuple[1])
        current_row_label = sheet_utils.format_addr(current_row_tuple)

        for i in range(0, self.currencies_len):
            currency_start = current_row_label
            currency_end = sheet_utils.format_addr((current_row_tuple[0] + 2, current_row_tuple[1]))
            currency_range = pygsheets.DataRange(start=currency_start, end=currency_end, worksheet=self.sheet)
            currency_range.merge_cells()

            current_row_label = sheet_utils.format_addr((current_row_tuple[0] + 3, current_row_tuple[1]))
            current_row_tuple = sheet_utils.format_addr(current_row_label)

        current_row_tuple = (self.current_exch_tuple[0] + 1, self.current_exch_tuple[1] + 1)
        current_row_label = sheet_utils.format_addr(current_row_tuple)

        current_cell = pygsheets.Cell(current_row_label, worksheet=self.sheet)
        current_cell.color = self.currency_info_rows_color
        current_cell.set_text_alignment(alignment="MIDDLE")
        current_cell.set_text_alignment(alignment="CENTER")

        range_start = current_row_label
        range_end = sheet_utils.format_addr((current_row_tuple[0] + self.currencies_len * 3 - 1, current_row_tuple[1]))
        currencies_range = pygsheets.DataRange(start=range_start, end=range_end, worksheet=self.sheet)
        currencies_range.applay_format(current_cell)

        time.sleep(1)

        # currency_info_rows_cell_list = []
        for i in range(0, self.currencies_len):

            for ii in range(0, len(self.currency_info_rows)):
                current_cell_tuple = (current_row_tuple[0] + ii, current_row_tuple[1])
                current_cell_label = sheet_utils.format_addr(current_cell_tuple)

                current_cell = pygsheets.Cell(current_cell_label, val=self.currency_info_rows[ii], worksheet=self.sheet)
                # current_cell.value = currency_info_rows[ii]
                exchange_update_cell_list.append(current_cell)

            current_row_tuple = (current_row_tuple[0] + len(self.currency_info_rows), current_row_tuple[1])
            current_row_label = sheet_utils.format_addr(current_row_label)

        # sheet.update_cells(cell_list=currency_info_rows_cell_list)

        # base_cell = pygsheets.Cell("A1", worksheet=sheet)
        # base_cell.set_text_alignment(alignment="MIDDLE")
        # base_cell.set_text_alignment(alignment="CENTER")
        # table_range.applay_format(base_cell)
        while True:
            count = 0
            try:
                self.sheet.update_cells(cell_list=exchange_update_cell_list)
                break
            except RequestError as e:
                log.warning("Encountered RequestError: {}".format(e))
                count += 1
                if count >= 10:
                    raise RequestError(e)

        time.sleep(10)

        self.current_exch_tuple = (self.current_exch_tuple[0] + 1 + self.currencies_len * 3 + self.row_spread,
                                   self.current_exch_tuple[1])

        self.current_exch_label = sheet_utils.format_addr(self.current_exch_tuple)

    def create_exchange_tables(self, force=False):

        self.sheet.adjust_column_width(start=self.start_cell_tuple[1] - 1, end=self.start_cell_tuple[1], pixel_size=90)
        self.sheet.adjust_column_width(start=self.start_cell_tuple[1], end=self.start_cell_tuple[1] + 26, pixel_size=75)



        # for exchange in self.exchanges:
        # log.warning(
        #     "\nself.current_exch_tuple: {}\nself.current_exch_label: {}".format(self.current_exch_tuple, self.current_exch_label))

    def test_function(self):
        pprint(self.sheet.find("Bitfinex")[0].label)
