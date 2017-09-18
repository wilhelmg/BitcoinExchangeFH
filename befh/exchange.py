#!/bin/python
from befh.zmq_client import ZmqClient
from befh.file_client import FileClient
from befh.market_data import L2Depth, Trade, Snapshot
from datetime import datetime
from threading import Lock


class ExchangeGateway:
    ############################################################################
    # Static variable 
    # Applied on all gateways whether to record the timestamp in local machine,
    # rather than exchange timestamp given by the API
    is_local_timestamp = True
    ############################################################################
    
    """
    Exchange gateway
    """
    def __init__(self, 
                 api_socket, 
                 db_clients=None):
        """
        Constructor
        :param exchange_name: Exchange name
        :param exchange_api: Exchange API
        :param db_client: Database client
        """
        self.db_clients = db_clients if db_clients else None
        self.api_socket = api_socket
        self.lock = Lock()
        self.exch_snapshot_id = 0

    @classmethod
    def get_exchange_name(cls):
        """
        Get exchange name
        :return: Exchange name string
        """
        return ''

    @classmethod
    def get_instmt_snapshot_table_name(cls, exchange, instmt_name):
        """
        Get instmt snapshot
        :param exchange: Exchange name
        :param instmt_name: Instrument name
        """
        return 'exch_' + exchange.lower() + '_' + instmt_name.lower() + \
               '_snapshot_' + datetime.utcnow().strftime("%Y%m%d")
        
    @classmethod
    def get_snapshot_table_name(cls):
        return 'exchanges_snapshot'

    @classmethod
    def is_allowed_snapshot(cls, db_client):
        return not isinstance(db_client, FileClient)

    @classmethod
    def is_allowed_instmt_record(cls, db_client):
        return not isinstance(db_client, ZmqClient)

    @classmethod
    def init_snapshot_table(cls, db_clients):
        for db_client in db_clients:
            db_client.create(cls.get_snapshot_table_name(),
                             Snapshot.columns(),
                             Snapshot.types(),
                             [0,1])
                             
    def init_instmt_snapshot_table(self, instmt):
        table_name = self.get_instmt_snapshot_table_name(instmt.get_exchange_name(),
                                                         instmt.get_instmt_name())
        for db_client in self.db_clients:
            db_client.create(table_name,
                             ['id'] + Snapshot.columns(False),
                             ['int'] + Snapshot.types(False),
                             [0])

    def start(self, instmt):
        """
        Start the exchange gateway
        :param instmt: Instrument
        :return List of threads
        """
        return []

    def get_instmt_snapshot_id(self, instmt):
        with self.lock:
            self.exch_snapshot_id += 1

        return self.exch_snapshot_id

    def insert_order_book(self, instmt):
        """
        Insert order book row into the database client
        :param instmt: Instrument
        """
        # If local timestamp indicator is on, assign the local timestamp again
        if self.is_local_timestamp:
            instmt.get_l2_depth().date_time = datetime.utcnow().strftime("%Y%m%d %H:%M:%S.%f")
        
        # Update the snapshot
        if instmt.get_l2_depth() is not None:
            id = self.get_instmt_snapshot_id(instmt)
            for db_client in self.db_clients:
                if self.is_allowed_snapshot(db_client):
                    db_client.insert(table=self.get_snapshot_table_name(),
                                     columns=Snapshot.columns(),
                                     types=Snapshot.types(),
                                     values=Snapshot.values(instmt.get_exchange_name(),
                                                            instmt.get_instmt_name(),
                                                            instmt.get_l2_depth(),
                                                            Trade() if instmt.get_last_trade() is None else instmt.get_last_trade(),
                                                            Snapshot.UpdateType.ORDER_BOOK),
                                     primary_key_index=[0,1],
                                     is_orreplace=True,
                                     is_commit=True)

                if self.is_allowed_instmt_record(db_client):
                    db_client.insert(table=instmt.get_instmt_snapshot_table_name(),
                                          columns=['id'] + Snapshot.columns(False),
                                          types=['int'] + Snapshot.types(False),
                                          values=[id] +
                                                  Snapshot.values('',
                                                                 '',
                                                                 instmt.get_l2_depth(),
                                                                 Trade() if instmt.get_last_trade() is None else instmt.get_last_trade(),
                                                                 Snapshot.UpdateType.ORDER_BOOK),
                                          is_commit=True)

    def insert_trade(self, instmt, trade):
        """
        Insert trade row into the database client
        :param instmt: Instrument
        """
        # If the instrument is not recovered, skip inserting into the table
        if not instmt.get_recovered():
            return
        
        # If local timestamp indicator is on, assign the local timestamp again
        if self.is_local_timestamp:
            trade.date_time = datetime.utcnow().strftime("%Y%m%d %H:%M:%S.%f")
        
        # Set the last trade to the current one
        instmt.set_last_trade(trade)

        # Update the snapshot
        if instmt.get_l2_depth() is not None and \
           instmt.get_last_trade() is not None:
            id = self.get_instmt_snapshot_id(instmt)
            for db_client in self.db_clients:
                is_allowed_snapshot = self.is_allowed_snapshot(db_client)
                is_allowed_instmt_record = self.is_allowed_instmt_record(db_client)
                if is_allowed_snapshot:
                    db_client.insert(table=self.get_snapshot_table_name(),
                                     columns=Snapshot.columns(),
                                     values=Snapshot.values(instmt.get_exchange_name(),
                                                            instmt.get_instmt_name(),
                                                            instmt.get_l2_depth(),
                                                            instmt.get_last_trade(),
                                                            Snapshot.UpdateType.TRADES),
                                     types=Snapshot.types(),
                                     primary_key_index=[0,1],
                                     is_orreplace=True,
                                     is_commit=not is_allowed_instmt_record)

                if is_allowed_instmt_record:
                    db_client.insert(table=instmt.get_instmt_snapshot_table_name(),
                                     columns=['id'] + Snapshot.columns(False),
                                     types=['int'] + Snapshot.types(False),
                                     values=[id] +
                                            Snapshot.values('',
                                                         '',
                                                         instmt.get_l2_depth(),
                                                         instmt.get_last_trade(),
                                                         Snapshot.UpdateType.TRADES),
                                     is_commit=True)
