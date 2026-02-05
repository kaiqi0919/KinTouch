# 出退勤管理システム - データベース管理モジュール

import sqlite3
import time
from datetime import datetime
from modules.constants import JST

class DatabaseManager:
    """データベース管理クラス"""
    
    def __init__(self, db_path):
        self.db_path = db_path
        self.init_database()
    
    def init_database(self):
        """データベース初期化"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # time_records テーブル（授業用）
            cursor.execute("PRAGMA table_info(time_records)")
            columns = cursor.fetchall()
            column_names = [col[1] for col in columns]
            
            if not columns:
                cursor.execute('''
                    CREATE TABLE time_records (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        instructor_id INTEGER,
                        card_uid TEXT NOT NULL,
                        instructor_name TEXT NOT NULL,
                        record_type TEXT NOT NULL CHECK (record_type IN ('IN', 'OUT')),
                        timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                ''')
            else:
                if 'instructor_name' not in column_names:
                    cursor.execute("ALTER TABLE time_records ADD COLUMN instructor_name TEXT DEFAULT '未登録'")
                if 'instructor_id' not in column_names:
                    cursor.execute("ALTER TABLE time_records ADD COLUMN instructor_id INTEGER")
            
            # meeting_records テーブル（会議用）
            cursor.execute("PRAGMA table_info(meeting_records)")
            meeting_columns = cursor.fetchall()
            
            if not meeting_columns:
                cursor.execute('''
                    CREATE TABLE meeting_records (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        instructor_id INTEGER,
                        card_uid TEXT NOT NULL,
                        instructor_name TEXT NOT NULL,
                        record_type TEXT NOT NULL CHECK (record_type IN ('IN', 'OUT')),
                        timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                ''')
            
            # instructors テーブル（講師マスタ）
            cursor.execute("PRAGMA table_info(instructors)")
            instructors_columns = cursor.fetchall()
            
            if not instructors_columns:
                cursor.execute('''
                    CREATE TABLE instructors (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        instructor_id INTEGER UNIQUE NOT NULL,
                        card_uid TEXT UNIQUE NOT NULL,
                        name TEXT NOT NULL,
                        created_at TIMESTAMP NOT NULL
                    )
                ''')
            
            conn.commit()
            conn.close()
            
        except Exception as e:
            print(f"データベース初期化エラー: {e}")
    
    def load_instructors(self):
        """DBから講師データ読み込み（UID→名前の辞書）"""
        instructors = {}
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute("SELECT card_uid, name FROM instructors")
            for card_uid, name in cursor.fetchall():
                instructors[card_uid] = name
            
            conn.close()
            return instructors
        except Exception as e:
            print(f"講師データ読み込みエラー: {e}")
            return {}
    
    def load_instructors_full(self):
        """DBから講師データ読み込み（全情報）"""
        instructors = []
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT instructor_id, card_uid, name, created_at
                FROM instructors
                ORDER BY instructor_id
            """)
            
            for instructor_id, card_uid, name, created_at in cursor.fetchall():
                instructors.append({
                    'instructor_id': str(instructor_id),
                    'card_uid': card_uid,
                    'name': name,
                    'created_at': created_at
                })
            
            conn.close()
            return instructors
        except Exception as e:
            print(f"講師データ読み込みエラー: {e}")
            return []
    
    def get_next_instructor_id(self):
        """次の講師番号を取得"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute("SELECT MAX(instructor_id) FROM instructors")
            result = cursor.fetchone()
            conn.close()
            
            max_id = result[0] if result[0] is not None else 0
            return max_id + 1
        except Exception as e:
            print(f"講師番号取得エラー: {e}")
            return 1
    
    def get_instructor_info_by_uid(self, card_uid):
        """UIDから講師情報取得"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT instructor_id, card_uid, name
                FROM instructors
                WHERE card_uid = ?
            """, (card_uid,))
            
            result = cursor.fetchone()
            conn.close()
            
            if result:
                return {
                    'instructor_id': result[0],
                    'card_uid': result[1],
                    'name': result[2]
                }
            return None
        except Exception as e:
            print(f"講師情報取得エラー: {e}")
            return None
    
    def get_instructor_info_by_id(self, instructor_id):
        """講師番号から講師情報取得"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT instructor_id, card_uid, name
                FROM instructors
                WHERE instructor_id = ?
            """, (instructor_id,))
            
            result = cursor.fetchone()
            conn.close()
            
            if result:
                return {
                    'instructor_id': str(result[0]),
                    'card_uid': result[1],
                    'name': result[2]
                }
            return None
        except Exception as e:
            print(f"講師情報取得エラー: {e}")
            return None
    
    def add_instructor_with_id(self, instructor_id, card_uid, name):
        """講師を指定IDで追加"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 重複チェック
            cursor.execute("SELECT instructor_id FROM instructors WHERE card_uid = ?", (card_uid,))
            if cursor.fetchone():
                conn.close()
                return False
            
            # 挿入
            created_at = datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")
            cursor.execute("""
                INSERT INTO instructors (instructor_id, card_uid, name, created_at)
                VALUES (?, ?, ?, ?)
            """, (instructor_id, card_uid, name, created_at))
            
            conn.commit()
            conn.close()
            return True
            
        except Exception as e:
            print(f"講師登録エラー: {e}")
            if 'conn' in locals():
                conn.close()
            return False
    
    def get_last_record(self, card_uid, table_name="time_records"):
        """最後の打刻記録を取得"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            query = f"SELECT record_type, timestamp FROM {table_name} WHERE card_uid = ? ORDER BY timestamp DESC LIMIT 1"
            cursor.execute(query, (card_uid,))
            
            result = cursor.fetchone()
            conn.close()
            
            if result:
                return {"type": result[0], "timestamp": result[1]}
            return None
            
        except Exception as e:
            print(f"最後の記録取得エラー: {e}")
            return None
    
    def record_attendance_to_db(self, card_uid, name, instructor_id, record_type, timestamp, table_name="time_records"):
        """データベースに打刻記録"""
        max_retries = 3
        for attempt in range(max_retries):
            try:
                conn = sqlite3.connect(self.db_path, timeout=10.0)
                cursor = conn.cursor()
                
                query = f"INSERT INTO {table_name} (instructor_id, card_uid, instructor_name, record_type, timestamp) VALUES (?, ?, ?, ?, ?)"
                cursor.execute(query, (instructor_id, card_uid, name, record_type, timestamp))
                
                conn.commit()
                conn.close()
                return True
                
            except sqlite3.OperationalError as e:
                if "database is locked" in str(e) and attempt < max_retries - 1:
                    time.sleep(0.5)
                    continue
                else:
                    print(f"打刻記録エラー: {e}")
                    return False
            except Exception as e:
                print(f"打刻記録エラー: {e}")
                return False
            finally:
                try:
                    if 'conn' in locals():
                        conn.close()
                except:
                    pass
        
        return False
    
    def get_date_records(self, date_str, table_name="time_records"):
        """特定日付の打刻記録取得"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            query = f'''
                SELECT instructor_name, record_type, timestamp 
                FROM {table_name}
                WHERE DATE(timestamp) = ?
                ORDER BY timestamp DESC
            '''
            cursor.execute(query, (date_str,))
            
            results = cursor.fetchall()
            conn.close()
            return results
            
        except Exception as e:
            print(f"記録取得エラー: {e}")
            return []
    
    def get_date_summary(self, date_str, table_name="time_records"):
        """特定日付のサマリー取得"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            query = f'''
                SELECT instructor_name, record_type, timestamp
                FROM {table_name}
                WHERE DATE(timestamp) = ?
                ORDER BY instructor_name, timestamp
            '''
            cursor.execute(query, (date_str,))
            
            results = cursor.fetchall()
            conn.close()
            
            if not results:
                return []
            
            # 講師ごとにグループ化
            instructor_records = {}
            for name, record_type, timestamp in results:
                if name not in instructor_records:
                    instructor_records[name] = []
                instructor_records[name].append((record_type, timestamp))
            
            summary = []
            for name, records in instructor_records.items():
                last_record = records[-1]
                status = "出勤中" if last_record[0] == "IN" else "退勤済"
                last_time = last_record[1]
                
                # その日の記録を文字列化
                record_str = ""
                for record_type, timestamp in records:
                    time_only = timestamp.split()[1][:5] if ' ' in timestamp else timestamp[:5]
                    action = "出" if record_type == "IN" else "退"
                    record_str += f"{action}:{time_only} "
                
                summary.append((name, status, last_time, record_str.strip()))
            
            return summary
            
        except Exception as e:
            print(f"サマリー取得エラー: {e}")
            return []
    
    def get_monthly_dates(self, month_str, table_name="time_records"):
        """対象月の日付一覧を取得"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            query = f'''
                SELECT DISTINCT DATE(timestamp) as date
                FROM {table_name}
                WHERE strftime('%Y-%m', timestamp) = ?
                ORDER BY date
            '''
            cursor.execute(query, (month_str,))
            dates = [row[0] for row in cursor.fetchall()]
            
            conn.close()
            return dates
        except Exception as e:
            print(f"日付一覧取得エラー: {e}")
            return []
    
    def get_monthly_summary_data(self, month_str, table_name="time_records"):
        """月次集計データ取得"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            query = f'''
                SELECT instructor_id, instructor_name, DATE(timestamp) as date
                FROM {table_name}
                WHERE strftime('%Y-%m', timestamp) = ?
                GROUP BY instructor_id, instructor_name, DATE(timestamp)
                ORDER BY instructor_id
            '''
            cursor.execute(query, (month_str,))
            
            results = cursor.fetchall()
            conn.close()
            return results
            
        except Exception as e:
            print(f"月次集計データ取得エラー: {e}")
            return []
    
    def get_instructor_monthly_records(self, month_str, instructor_id, table_name="time_records"):
        """講師の月次打刻記録を取得"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            query = f'''
                SELECT DATE(timestamp) as date, TIME(timestamp) as time
                FROM {table_name}
                WHERE instructor_id = ? AND strftime('%Y-%m', timestamp) = ?
                ORDER BY date, time
            '''
            cursor.execute(query, (instructor_id, month_str))
            records = cursor.fetchall()
            
            conn.close()
            return records
        except Exception as e:
            print(f"講師別記録取得エラー: {e}")
            return []