# 簡易版継続監視用カードリーダーコード（監視オプションなし）
# pip install pyscard

import time
from smartcard.System import readers
from smartcard.util import toHexString
from smartcard.Exceptions import CardConnectionException, NoCardException

class SimpleCardMonitor:
    def __init__(self):
        self.reader = None
        self.connection = None
        self.last_uid = None
        
    def initialize_reader(self):
        """リーダーを初期化"""
        try:
            r = readers()
            if len(r) == 0:
                print("カードリーダーが見つかりません。")
                return False
            
            print("利用可能なリーダー:")
            for i, reader in enumerate(r):
                print(f"{i}: {reader}")
            
            self.reader = r[0]
            print(f"リーダー '{self.reader}' を使用します。")
            return True
            
        except Exception as e:
            print(f"リーダー初期化エラー: {e}")
            return False
    
    def connect_to_card(self):
        """カードに接続（毎回新しい接続を作成）"""
        try:
            # 既存の接続があれば切断
            if self.connection:
                try:
                    self.connection.disconnect()
                except:
                    pass
                self.connection = None
            
            # 新しい接続を作成
            self.connection = self.reader.createConnection()
            self.connection.connect()
            return True
            
        except NoCardException:
            # カードが存在しない場合は正常な状態
            return False
        except CardConnectionException as e:
            # 接続エラーは静かに処理（監視中は頻繁に発生する可能性）
            return False
        except Exception as e:
            print(f"予期しないエラー: {e}")
            return False
    
    def get_card_uid(self):
        """カードのUIDを取得"""
        if not self.connection:
            return None
            
        try:
            # UIDを取得
            get_uid = [0xFF, 0xCA, 0x00, 0x00, 0x00]
            data, sw1, sw2 = self.connection.transmit(get_uid)
            
            if sw1 == 0x90 and sw2 == 0x00:
                uid = toHexString(data)
                return uid
            else:
                return None
                
        except Exception as e:
            return None
    
    def disconnect(self):
        """接続を切断"""
        if self.connection:
            try:
                self.connection.disconnect()
            except:
                pass
            self.connection = None
    
    def is_card_present(self):
        """カードが存在するかチェック"""
        try:
            # 接続を試行してカードの存在を確認
            test_connection = self.reader.createConnection()
            test_connection.connect()
            test_connection.disconnect()
            return True
        except (NoCardException, CardConnectionException):
            return False
        except Exception:
            return False
    
    def monitor_cards(self):
        """継続的にカードを監視（UIDのみ表示、0.5秒間隔）"""
        print("カード監視を開始します（UIDのみ表示、0.5秒間隔）。Ctrl+Cで終了。")
        print("-" * 50)
        
        try:
            while True:
                if self.is_card_present():
                    if self.connect_to_card():
                        uid = self.get_card_uid()
                        if uid:
                            # 前回と異なるカードの場合のみ表示
                            if uid != self.last_uid:
                                print(f"新しいカード検出: UID = {uid}")
                                self.last_uid = uid
                        
                        # 処理後は必ず切断
                        self.disconnect()
                else:
                    # カードが離れた場合
                    if self.last_uid:
                        print("カードが離れました。")
                        self.last_uid = None
                
                time.sleep(0.5)  # 0.5秒間隔で監視
                
        except KeyboardInterrupt:
            print("\n監視を終了します。")
        finally:
            self.disconnect()

def main():
    """メイン関数"""
    card_reader = SimpleCardMonitor()
    
    if not card_reader.initialize_reader():
        return
    
    # 監視オプションなし、直接監視開始
    try:
        card_reader.monitor_cards()
            
    except KeyboardInterrupt:
        print("\nプログラムを終了します。")
    finally:
        card_reader.disconnect()

if __name__ == "__main__":
    main()
