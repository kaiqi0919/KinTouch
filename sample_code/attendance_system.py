# 出退勤管理システム
# pip install pyscard

import time
import sqlite3
from datetime import datetime
from smartcard.System import readers
from smartcard.util import toHexString
from smartcard.Exceptions import CardConnectionException, NoCardException

class AttendanceSystem:
    def __init__(self, db_path="attendance.db"):
        self.reader = None
        self.connection = None
        self.last_uid = None
        self.db_path = db_path
        self.init_database()
        
    def init_database(self):
        """データベースとテーブルを初期化"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 講師マスタテーブル
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS instructors (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    card_uid TEXT UNIQUE NOT NULL,
                    name TEXT NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            # 打刻時刻テーブル
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS time_records (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    instructor_id INTEGER NOT NULL,
                    card_uid TEXT NOT NULL,
                    record_type TEXT NOT NULL CHECK (record_type IN ('IN', 'OUT')),
                    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (instructor_id) REFERENCES instructors (id)
                )
            ''')
            
            conn.commit()
            conn.close()
            print("データベースを初期化しました。")
            
        except Exception as e:
            print(f"データベース初期化エラー: {e}")
    
    def add_instructor(self, card_uid, name):
        """講師を登録"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute(
                "INSERT INTO instructors (card_uid, name) VALUES (?, ?)",
                (card_uid, name)
            )
            
            conn.commit()
            instructor_id = cursor.lastrowid
            conn.close()
            
            print(f"講師を登録しました: {name} (UID: {card_uid})")
            return instructor_id
            
        except sqlite3.IntegrityError:
            print(f"エラー: カードUID {card_uid} は既に登録されています。")
            return None
        except Exception as e:
            print(f"講師登録エラー: {e}")
            return None
    
    def get_instructor_by_uid(self, card_uid):
        """UIDから講師情報を取得"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute(
                "SELECT id, name FROM instructors WHERE card_uid = ?",
                (card_uid,)
            )
            
            result = cursor.fetchone()
            conn.close()
            
            if result:
                return {"id": result[0], "name": result[1]}
            return None
            
        except Exception as e:
            print(f"講師情報取得エラー: {e}")
            return None
    
    def get_last_record(self, instructor_id):
        """講師の最後の打刻記録を取得"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute(
                "SELECT record_type, timestamp FROM time_records WHERE instructor_id = ? ORDER BY timestamp DESC LIMIT 1",
                (instructor_id,)
            )
            
            result = cursor.fetchone()
            conn.close()
            
            if result:
                return {"type": result[0], "timestamp": result[1]}
            return None
            
        except Exception as e:
            print(f"最後の記録取得エラー: {e}")
            return None
    
    def record_attendance(self, card_uid):
        """出退勤を記録"""
        instructor = self.get_instructor_by_uid(card_uid)
        
        if not instructor:
            print(f"未登録のカードです (UID: {card_uid})")
            return False
        
        # 最後の記録を確認して出勤/退勤を判定
        last_record = self.get_last_record(instructor["id"])
        
        if last_record is None or last_record["type"] == "OUT":
            record_type = "IN"
            action = "出勤"
        else:
            record_type = "OUT"
            action = "退勤"
        
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute(
                "INSERT INTO time_records (instructor_id, card_uid, record_type) VALUES (?, ?, ?)",
                (instructor["id"], card_uid, record_type)
            )
            
            conn.commit()
            conn.close()
            
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f"【{action}】{instructor['name']} さん ({current_time})")
            return True
            
        except Exception as e:
            print(f"打刻記録エラー: {e}")
            return False
    
    def list_instructors(self):
        """登録済み講師一覧を表示"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute("SELECT card_uid, name, created_at FROM instructors ORDER BY name")
            results = cursor.fetchall()
            conn.close()
            
            if results:
                print("\n=== 登録済み講師一覧 ===")
                for uid, name, created_at in results:
                    print(f"名前: {name}, UID: {uid}, 登録日: {created_at}")
            else:
                print("登録済み講師はいません。")
                
        except Exception as e:
            print(f"講師一覧取得エラー: {e}")
    
    def show_today_records(self):
        """今日の打刻記録を表示"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT i.name, tr.record_type, tr.timestamp 
                FROM time_records tr
                JOIN instructors i ON tr.instructor_id = i.id
                WHERE DATE(tr.timestamp) = DATE('now', 'localtime')
                ORDER BY tr.timestamp DESC
            ''')
            
            results = cursor.fetchall()
            conn.close()
            
            if results:
                print("\n=== 今日の打刻記録 ===")
                for name, record_type, timestamp in results:
                    action = "出勤" if record_type == "IN" else "退勤"
                    print(f"{timestamp} - {name} さん ({action})")
            else:
                print("今日の打刻記録はありません。")
                
        except Exception as e:
            print(f"今日の記録取得エラー: {e}")
    
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
        """カードに接続"""
        try:
            if self.connection:
                try:
                    self.connection.disconnect()
                except:
                    pass
                self.connection = None
            
            self.connection = self.reader.createConnection()
            self.connection.connect()
            return True
            
        except NoCardException:
            return False
        except CardConnectionException:
            return False
        except Exception as e:
            print(f"予期しないエラー: {e}")
            return False
    
    def get_card_uid(self):
        """カードのUIDを取得"""
        if not self.connection:
            return None
            
        try:
            get_uid = [0xFF, 0xCA, 0x00, 0x00, 0x00]
            data, sw1, sw2 = self.connection.transmit(get_uid)
            
            if sw1 == 0x90 and sw2 == 0x00:
                uid = toHexString(data)
                return uid
            else:
                return None
                
        except Exception:
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
            test_connection = self.reader.createConnection()
            test_connection.connect()
            test_connection.disconnect()
            return True
        except (NoCardException, CardConnectionException):
            return False
        except Exception:
            return False
    
    def monitor_attendance(self):
        """出退勤監視を開始"""
        print("出退勤監視を開始します。Ctrl+Cで終了。")
        print("-" * 50)
        
        try:
            while True:
                if self.is_card_present():
                    if self.connect_to_card():
                        uid = self.get_card_uid()
                        if uid and uid != self.last_uid:
                            self.record_attendance(uid)
                            self.last_uid = uid
                        
                        self.disconnect()
                else:
                    if self.last_uid:
                        self.last_uid = None
                
                time.sleep(0.5)
                
        except KeyboardInterrupt:
            print("\n監視を終了します。")
        finally:
            self.disconnect()

def main():
    """メイン関数"""
    system = AttendanceSystem()
    
    if not system.initialize_reader():
        return
    
    while True:
        print("\n=== 出退勤管理システム ===")
        print("1. 出退勤監視開始")
        print("2. 講師登録")
        print("3. 講師一覧表示")
        print("4. 今日の打刻記録表示")
        print("5. 終了")
        
        choice = input("選択してください (1-5): ").strip()
        
        if choice == "1":
            try:
                system.monitor_attendance()
            except KeyboardInterrupt:
                print("\n監視を終了しました。")
                
        elif choice == "2":
            print("\n講師登録のため、カードをリーダーにかざしてください...")
            
            # カード読み取り待機
            card_detected = False
            while not card_detected:
                if system.is_card_present():
                    if system.connect_to_card():
                        uid = system.get_card_uid()
                        if uid:
                            print(f"カードを検出しました (UID: {uid})")
                            name = input("講師名を入力してください: ").strip()
                            if name:
                                system.add_instructor(uid, name)
                            card_detected = True
                        system.disconnect()
                time.sleep(0.5)
                
        elif choice == "3":
            system.list_instructors()
            
        elif choice == "4":
            system.show_today_records()
            
        elif choice == "5":
            print("システムを終了します。")
            break
            
        else:
            print("無効な選択です。")

if __name__ == "__main__":
    main()
