import os
import sys
import json
import pygsheets

from pygsheets.exceptions import RequestError

PROJECT_ROOT = os.path.abspath(os.path.dirname(__file__))


def update_doc(exchange, price):
    while True:
        count = 0
        try:
            exchange_cell = get_cell_number(exchange)

            gc = auth_sheets()

            doc = gc.open_by_key("1N1MDj7mXLbv_-LUj1peug53P9a2gh5jdD5BnPrOQVLU")

            sheet = doc.worksheet_by_title("ArbitrageTable")

            sheet.update_cell(exchange_cell, price)
            return True

        except RequestError as e:
            print("Encountered Request error: %s" % e)
            count += 1

        if count > 10:
            return False


def get_cell_number(exchange):
    cells_config = PROJECT_ROOT + os.sep + "exchange_cells.json"
    with open(cells_config, 'r') as read:
        cells = json.load(read)

    return cells[exchange]['price']


def auth_sheets(path_to_creds=None):
    if not path_to_creds:
        if sys.platform == "win32":
            credentials = "C:\\Users\\Wilhelm\\PycharmProjects\\arbitrage\\config\\service_creds.json"
        else:
            credentials = "/home/guldenmw/PycharmProjects/arbitrage/config/service_creds.json"
    else:
        credentials = path_to_creds

    return pygsheets.authorize(service_file=credentials, no_cache=True)
