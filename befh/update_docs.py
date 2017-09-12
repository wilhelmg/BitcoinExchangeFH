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

logging.basicConfig(level=logging.DEBUG)
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

    def track_request_limit(self):
        start =

    def get_instmt_number(self, exchange, instmt):
        exch_table_tuples = self.get_table_tuples(exchange)
        if not exch_table_tuples:
            log.warning("Failed to determine %s table cells." % exchange)
            return

        if exchange not in self.sheet_info['exchanges'] \
                or instmt not in self.sheet_info['exchanges'][exchange]['currency_pairs']:
            log.info("%s not found in %s." % instmt, exchange)
            return

        base = self.sheet_info['exchanges'][exchange]['currency_pairs'][instmt]['base']
        base_type = "crypto"
        counter = self.sheet_info['exchanges'][exchange]['currency_pairs'][instmt]['counter']
        counter_type = "fiat"

        if base != "BTC" and base in self.fiat:
            base_type = "fiat"

        if counter in self.currencies:
            counter_type = "crypto"

        if base_type == "crypto" and counter_type == "crypto" and base != "USDT":
            if base == "BTC":
                base, counter = counter, base

        if base == "USDT":
            base, counter = counter, base[:-1]

        log.debug("\nbase: {}\nbase_type: {}\ncounter: {}\ncounter_type: {}\n".format(base, base_type,
                                                                                      counter, counter_type))

        currency_rows = self.get_currency_rows(exchange, returnas='cell')
        fiat_columns = self.get_fiat_columns(exchange, returnas='cell')
        log.debug("\ncurrency_rows: \n{}\nfiat_columns: \n{}\n".format(currency_rows, fiat_columns))

        currency_cell = None
        for cell in currency_rows:
            # log.debug("cell.value: {}".format(cell.value))
            if cell.value == base:
                log.debug("{} == {}".format(cell.value, base))
                currency_cell = cell

        fiat_cell = None
        for cell in fiat_columns:
            # log.debug("cell.value: {}".format(cell.value))
            if cell.value == counter:
                log.debug("{} == {}".format(cell.value, counter))
                fiat_cell = cell

        return currency_cell.row, fiat_cell.col

    def get_fiat_columns(self, exchange, returnas='matrix'):
        exch_table_labels = self.get_table_labels(exchange)
        columns = self.sheet.get_values(start=exch_table_labels[0],
                                       end=exch_table_labels[2],
                                       returnas=returnas,
                                       include_empty=False)
        if not columns:
            log.warning("Could not find values for requested columns: {}".format(columns))
            return
        elif len(columns) > 1:
            log.warning("\nMore than one value present in columns: \n{}\n".format(pformat(columns)))
            return
        return columns[0]

    def get_currency_rows(self, exchange, returnas='matrix'):
        exch_table_labels = self.get_table_labels(exchange)
        raw_rows = self.sheet.get_values(start=exch_table_labels[0],
                                         end=exch_table_labels[1],
                                         returnas=returnas,
                                         include_empty=False)
        log.debug("\nraw_rows: \n{}\n".format(raw_rows))

        rows = [i[0] for i in raw_rows if i]
        if not rows:
            log.warning("Could not find values for requested rows: {}".format(rows))
            return
        return rows

    def find_exchange_label(self, exchange):
        exchange_start_cell = self.sheet.find(query=exchange)
        if not exchange_start_cell or len(exchange_start_cell) > 2:
            log.warning("Failed to find exchange labels. \nStart cell: {}".format(exchange_start_cell))
            return False

        log.debug("Start cell: {}".format(exchange_start_cell[0].label))
        return exchange_start_cell[0].label

    def get_table_labels(self, exchange=None, tuples=None):
        if not exchange and not tuple:
            raise ValueError("Please provide either exchange name or tuples.")
        if not tuples:
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

        lt = sheet_utils.format_addr(start_label)
        lb = (lt[0] + self.currencies_len * 3, lt[1])
        rt = (lt[0], lt[1] + 1 + self.fiat_len + len(self.info_columns))
        rb = (lb[0], rt[1])

        return lt, lb, rt, rb

    def validate_exch_tables(self):
        missing = []
        invalid = []
        valid = []
        updated = []
        for exchange in self.exchanges:
            status = self.validate_exch_table(exchange)
            if status == "Missing":
                missing.append(exchange)
            elif status == "Invalid":
                invalid.append(exchange)
            elif status == "Valid":
                valid.append(exchange)
            else:
                raise ValueError("Invalid status: {}.".format(status))

        if missing:
            log.warning("\nDetected missing exchange tables. Rebuilding all tables. \n{}\n.".format(pformat(missing)))
            self.create_exchange_tables()
            return
        if invalid:
            log.warning("\nDetected invalid exchange tables. Rebuilding all tables. \n{}\n.".format(pformat(invalid)))
            for exchange in invalid:
                self.create_exchange_table(exchange)
                updated.append(exchange)
        if valid:
            log.debug("\nSuccessfully validated following tables:\n{}\n".format(pformat(valid)))
            log.debug("\nSuccessfully updated following tables:\n{}\n".format(pformat(updated)))
            return True

    def validate_exch_table(self, exchange):
        log.debug("Validating %s table..." % exchange)
        table_tuples = self.get_table_tuples(exchange)
        if not table_tuples:
            return "Missing"
        lt, lb, rt, rb = table_tuples[0], table_tuples[1], table_tuples[2], table_tuples[3]
        log.debug("{} table starting cell: {}".format(exchange, lt))
        columns = self.get_fiat_columns(exchange=exchange)[0]
        rows = self.get_currency_rows(exchange=exchange)
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

    def create_exchange_tables(self):
        self.sheet.adjust_column_width(start=self.start_cell_tuple[1] - 1, end=self.start_cell_tuple[1], pixel_size=90)
        self.sheet.adjust_column_width(start=self.start_cell_tuple[1], end=self.start_cell_tuple[1] + 26, pixel_size=75)

        self.current_exch_label = self.start_cell_label
        self.current_exch_tuple = self.start_cell_tuple

        for exchange in self.exchanges:
            while True:
                count = 1
                try:
                    self.create_exchange_table(exchange)
                    log.debug("Created %s table successfully." % exchange)
                    break
                except RequestError as e:
                    log.warning("Encountered RequestError: {}. Retrying...".format(e))
                    if count > 10:
                        raise RequestError("Encountered more than 10 RequestErrors. Terminating")
                    count += 1
            time.sleep(10)

    def create_exchange_table(self, exchange):
        log.info("Creating %s table..." % exchange)
        exchange_update_cell_list = []
        exchange_info = self.sheet_info['exchanges'][self.exchanges.index(exchange)][exchange]

        exchange_cell = pygsheets.Cell(self.current_exch_label, val=exchange, worksheet=self.sheet)
        # exchange_update_cell_list.append(exchange_cell)
        exchange_cell.set_text_alignment(alignment="MIDDLE")
        exchange_cell.set_text_alignment(alignment="CENTER")

        lt_tuple, lb_tuple, rt_tuple, rb_tuple = self.get_table_tuples(exchange)
        lt_label, lb_label, rt_label, rb_label = self.get_table_labels(tuples=(lt_tuple, lb_tuple, rt_tuple, rb_tuple))
        log.debug("lt_label: {}\nlb_label: {}\nrt_label: {}\nrb_label: {}\n".format(lt_label,
                                                                                    lb_label,
                                                                                    rt_label,
                                                                                    rb_label))

        table_range = pygsheets.DataRange(start=lt_label, end=rb_label, worksheet=self.sheet)
        exchange_cell.color = tuple(exchange_info['info']['cell_color'])

        table_range.update_borders()

        # exchange_cell.value = exchange

        current_col_tuple = (lt_tuple[0], lt_tuple[1] + 1)
        current_col_label = sheet_utils.format_addr(current_col_tuple)

        current_col_cell = pygsheets.Cell(current_col_label, val="Category", worksheet=self.sheet)

        current_col_cell.color = self.categories_color
        current_col_cell.set_text_alignment(alignment="MIDDLE")
        current_col_cell.set_text_alignment(alignment="CENTER")

        column_range = pygsheets.DataRange(
            start=current_col_label,
            end=sheet_utils.format_addr((current_col_tuple[0], rt_tuple[1])),
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


        # for exchange in self.exchanges:
        # log.warning(
        #     "\nself.current_exch_tuple: {}\nself.current_exch_label: {}".format(self.current_exch_tuple, self.current_exch_label))

    def update_trade_cell(self, exchange, instmt, price):
        log.debug("\nexchange: {}\ninstmt: {}\nprice: {}".format(exchange, instmt, price))
        while True:
            count = 0
            try:
                price_tuple = self.get_instmt_number(exchange=exchange, instmt=instmt)
                # price_label = sheet_utils.format_addr(price_tuple)
                if not price_tuple:
                    raise ValueError("Could not locate instmt cell.")

                self.sheet.update_cell(price_tuple, price)
                return True

            except RequestError as e:
                print("Encountered Request error: %s" % e)
                if count > 10:
                    return False
                count += 1

    def test_function(self):
        pprint(self.sheet.find("Bitfinex")[0].label)
