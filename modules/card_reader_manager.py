# 出退勤管理システム - カードリーダー管理モジュール

from smartcard.System import readers
from smartcard.util import toHexString
from smartcard.Exceptions import CardConnectionException, NoCardException


class CardReaderManager:
    """カードリーダー管理クラス"""
    
    def __init__(self):
        self.class_reader = None
        self.meeting_reader = None
        self.class_reader_name = None
        self.meeting_reader_name = None
    
    def get_available_readers(self):
        """利用可能なリーダーのリストを取得"""
        try:
            r = readers()
            return [reader.name for reader in r]
        except Exception as e:
            print(f"リーダー取得エラー: {e}")
            return []
    
    def initialize_readers(self, class_reader_name, meeting_reader_name):
        """リーダー初期化"""
        try:
            r = readers()
            if len(r) < 2:
                return False
            
            self.class_reader_name = class_reader_name
            self.meeting_reader_name = meeting_reader_name
            
            # リーダーを名前で識別
            for reader in r:
                if reader.name == class_reader_name:
                    self.class_reader = reader
                elif reader.name == meeting_reader_name:
                    self.meeting_reader = reader
            
            if self.class_reader is None or self.meeting_reader is None:
                return False
            
            return True
            
        except Exception as e:
            print(f"リーダー初期化エラー: {e}")
            return False
    
    def connect_to_card(self, reader):
        """カードに接続"""
        try:
            connection = reader.createConnection()
            connection.connect()
            return connection
        except (NoCardException, CardConnectionException):
            return None
        except Exception:
            return None
    
    def get_card_uid(self, connection):
        """カードUID取得"""
        if not connection:
            return None
            
        try:
            get_uid = [0xFF, 0xCA, 0x00, 0x00, 0x00]
            data, sw1, sw2 = connection.transmit(get_uid)
            
            if sw1 == 0x90 and sw2 == 0x00:
                uid = toHexString(data)
                return uid
            else:
                return None
                
        except Exception:
            return None
    
    def disconnect(self, connection):
        """接続切断"""
        if connection:
            try:
                connection.disconnect()
            except:
                pass
    
    def is_card_present(self, reader):
        """カード存在チェック"""
        try:
            test_connection = reader.createConnection()
            test_connection.connect()
            test_connection.disconnect()
            return True
        except (NoCardException, CardConnectionException):
            return False
        except Exception:
            return False