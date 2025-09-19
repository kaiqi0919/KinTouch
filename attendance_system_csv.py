# 出退勤管理システム（講師マスタCSV版）
# pip install pyscard

import time
import sqlite3
import csv
import os
import winsound
from datetime import datetime, timezone, timedelta
from smartcard.System import readers
from smartcard.util import toHexString
from smartcard.Exceptions import CardConnectionException, NoCardException

# 日本時間のタイムゾーン設定
JST = timezone(timedelta(hours=9))

class AttendanceSystemCSV:
    def __init__(self, db_path="attendance.db", instructors_csv="instructors.csv"):
        self.reader = None
        self.connection = None
        self.last_uid = None
        self.db_path = db_path
        self.instructors_csv = instructors_csv
        self.sound_enabled = True  # 音声機能の有効/無効
        self.init_database()
        self.init_instructors_csv()
    
    def play_beep(self, beep_type="success"):
        """ビープ音を再生"""
        if not self.sound_enabled:
            return
        
        try:
            if beep_type == "success":
                # 成功音：高い音で短く2回
                winsound.Beep(1000, 200)  # 1000Hz, 200ms
                time.sleep(0.1)
                winsound.Beep(1000, 200)
            elif beep_type == "error":
                # エラー音：低い音で長く1回
                winsound.Beep(400, 500)   # 400Hz, 500ms
            elif beep_type == "card_detected":
                # カード検出音：中程度の音で短く1回
                winsound.Beep(800, 150)   # 800Hz, 150ms
        except Exception as e:
            print(f"音声再生エラー: {e}")
    
    def toggle_sound(self):
        """音声機能のオン/オフを切り替え"""
        self.sound_enabled = not self.sound_enabled
        status = "有効" if self.sound_enabled else "無効"
        print(f"音声機能を{status}にしました。")
        
        # テスト音を再生
        if self.sound_enabled:
            self.play_beep("success")
        
    def init_database(self):
        """データベースとテーブルを初期化（打刻記録のみ）"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 既存テーブルの構造を確認
            cursor.execute("PRAGMA table_info(time_records)")
            columns = cursor.fetchall()
            column_names = [col[1] for col in columns]
            
            # instructor_nameカラムが存在しない場合は追加
            if 'instructor_name' not in column_names and columns:
                print("既存テーブルにinstructor_nameカラムを追加します...")
                cursor.execute("ALTER TABLE time_records ADD COLUMN instructor_name TEXT DEFAULT '未登録'")
                
                # 既存データのinstructor_nameをデフォルト値で更新
                cursor.execute("UPDATE time_records SET instructor_name = '未登録' WHERE instructor_name IS NULL")
            
            # テーブルが存在しない場合は新規作成
            elif not columns:
                cursor.execute('''
                    CREATE TABLE time_records (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        card_uid TEXT NOT NULL,
                        instructor_name TEXT NOT NULL,
                        record_type TEXT NOT NULL CHECK (record_type IN ('IN', 'OUT')),
                        timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                ''')
            
            conn.commit()
            conn.close()
            print("データベースを初期化しました。")
            
        except Exception as e:
            print(f"データベース初期化エラー: {e}")
    
    def init_instructors_csv(self):
        """講師マスタCSVファイルを初期化"""
        if not os.path.exists(self.instructors_csv):
            try:
                with open(self.instructors_csv, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(["instructor_id", "card_uid", "name", "created_at"])
                print(f"講師マスタCSVファイル '{self.instructors_csv}' を作成しました。")
            except Exception as e:
                print(f"講師マスタCSV初期化エラー: {e}")
    
    def load_instructors(self):
        """CSVファイルから講師データを読み込み"""
        instructors = {}
        try:
            if os.path.exists(self.instructors_csv):
                with open(self.instructors_csv, 'r', encoding='utf-8') as csvfile:
                    reader = csv.DictReader(csvfile)
                    for row in reader:
                        instructors[row['card_uid']] = row['name']
            return instructors
        except Exception as e:
            print(f"講師データ読み込みエラー: {e}")
            return {}
    
    def get_next_instructor_id(self):
        """次の講師番号を取得"""
        try:
            if not os.path.exists(self.instructors_csv):
                return 1
            
            with open(self.instructors_csv, 'r', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile)
                max_id = 0
                for row in reader:
                    try:
                        instructor_id = int(row['instructor_id'])
                        max_id = max(max_id, instructor_id)
                    except (ValueError, KeyError):
                        continue
                return max_id + 1
        except Exception as e:
            print(f"講師番号取得エラー: {e}")
            return 1
    
    def add_instructor(self, card_uid, name):
        """講師をCSVファイルに追加"""
        try:
            # 既存データをチェック
            instructors = self.load_instructors()
            if card_uid in instructors:
                print(f"エラー: カードUID {card_uid} は既に登録されています。")
                return False
            
            # 次の講師番号を取得
            instructor_id = self.get_next_instructor_id()
            
            # CSVファイルに追加
            with open(self.instructors_csv, 'a', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                created_at = datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")
                writer.writerow([instructor_id, card_uid, name, created_at])
            
            print(f"講師を登録しました: {name} (講師番号: {instructor_id}, UID: {card_uid})")
            return True
            
        except Exception as e:
            print(f"講師登録エラー: {e}")
            return False
    
    def get_instructor_by_uid(self, card_uid):
        """UIDから講師情報を取得"""
        instructors = self.load_instructors()
        return instructors.get(card_uid)
    
    def get_last_record(self, card_uid):
        """講師の最後の打刻記録を取得"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute(
                "SELECT record_type, timestamp FROM time_records WHERE card_uid = ? ORDER BY timestamp DESC LIMIT 1",
                (card_uid,)
            )
            
            result = cursor.fetchone()
            conn.close()
            
            if result:
                return {"type": result[0], "timestamp": result[1]}
            return None
            
        except Exception as e:
            print(f"最後の記録取得エラー: {e}")
            return None
    
    def get_instructor_info_by_uid(self, card_uid):
        """UIDから講師情報（ID、名前）を取得"""
        try:
            if not os.path.exists(self.instructors_csv):
                return None
            
            with open(self.instructors_csv, 'r', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    if row['card_uid'] == card_uid:
                        return {
                            'instructor_id': row['instructor_id'],
                            'name': row['name']
                        }
            return None
        except Exception as e:
            print(f"講師情報取得エラー: {e}")
            return None
    
    def record_attendance(self, card_uid):
        """出退勤を記録"""
        instructor_info = self.get_instructor_info_by_uid(card_uid)
        
        if not instructor_info:
            print(f"未登録のカードです (UID: {card_uid})")
            # エラー音を再生
            self.play_beep("error")
            return False
        
        instructor_name = instructor_info['name']
        instructor_id = instructor_info['instructor_id']
        
        # 最後の記録を確認して出勤/退勤を判定
        last_record = self.get_last_record(card_uid)
        
        if last_record is None or last_record["type"] == "OUT":
            record_type = "IN"
            action = "出勤"
        else:
            record_type = "OUT"
            action = "退勤"
        
        # データベースロック対策：リトライ機能付き
        max_retries = 3
        for attempt in range(max_retries):
            try:
                conn = sqlite3.connect(self.db_path, timeout=10.0)
                cursor = conn.cursor()
                
                # テーブル構造を確認
                cursor.execute("PRAGMA table_info(time_records)")
                columns = cursor.fetchall()
                column_names = [col[1] for col in columns]
                
                # 日本時間で現在時刻を取得
                jst_now = datetime.now(JST)
                timestamp_str = jst_now.strftime("%Y-%m-%d %H:%M:%S")
                
                # instructor_idカラムが存在する場合
                if 'instructor_id' in column_names:
                    cursor.execute(
                        "INSERT INTO time_records (instructor_id, card_uid, instructor_name, record_type, timestamp) VALUES (?, ?, ?, ?, ?)",
                        (instructor_id, card_uid, instructor_name, record_type, timestamp_str)
                    )
                else:
                    cursor.execute(
                        "INSERT INTO time_records (card_uid, instructor_name, record_type, timestamp) VALUES (?, ?, ?, ?)",
                        (card_uid, instructor_name, record_type, timestamp_str)
                    )
                
                conn.commit()
                conn.close()
                
                print(f"【{action}】{instructor_name} さん ({timestamp_str})")
                # 成功音を再生
                self.play_beep("success")
                return True
                
            except sqlite3.OperationalError as e:
                if "database is locked" in str(e) and attempt < max_retries - 1:
                    print(f"データベースがロックされています。再試行中... ({attempt + 1}/{max_retries})")
                    time.sleep(0.5)  # 0.5秒待機してリトライ
                    continue
                else:
                    print(f"打刻記録エラー: {e}")
                    # エラー音を再生
                    self.play_beep("error")
                    return False
            except Exception as e:
                print(f"打刻記録エラー: {e}")
                # エラー音を再生
                self.play_beep("error")
                return False
            finally:
                try:
                    if 'conn' in locals():
                        conn.close()
                except:
                    pass
        
        # 最大リトライ回数に達した場合
        self.play_beep("error")
        return False
    
    def list_instructors(self):
        """登録済み講師一覧を表示"""
        try:
            if not os.path.exists(self.instructors_csv):
                print("講師マスタファイルが存在しません。")
                return
            
            with open(self.instructors_csv, 'r', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile)
                instructors = list(reader)
            
            if instructors:
                print(f"\n=== 登録済み講師一覧 ({len(instructors)}人) ===")
                print(f"{'講師番号':<8} {'カードUID':<20} {'講師名':<15} {'登録日時':<20}")
                print("-" * 70)
                for instructor in instructors:
                    instructor_id = instructor.get('instructor_id', 'N/A')
                    print(f"{instructor_id:<8} {instructor['card_uid']:<20} {instructor['name']:<15} {instructor['created_at']:<20}")
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
                SELECT instructor_name, record_type, timestamp 
                FROM time_records
                WHERE DATE(timestamp) = DATE('now', 'localtime')
                ORDER BY timestamp DESC
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
    
    def show_date_records(self):
        """特定の日付の打刻記録を表示"""
        print("\n=== 特定の日付の打刻記録表示 ===")
        
        # 日付入力
        while True:
            date_input = input("日付を入力してください (YYYY-MM-DD形式、例: 2025-01-15): ").strip()
            
            if not date_input:
                print("日付の入力をキャンセルしました。")
                return
            
            # 日付形式の検証
            try:
                # 日付の妥当性をチェック
                datetime.strptime(date_input, "%Y-%m-%d")
                break
            except ValueError:
                print("無効な日付形式です。YYYY-MM-DD形式で入力してください。")
                continue
        
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT instructor_name, record_type, timestamp 
                FROM time_records
                WHERE DATE(timestamp) = ?
                ORDER BY timestamp DESC
            ''', (date_input,))
            
            results = cursor.fetchall()
            conn.close()
            
            if results:
                print(f"\n=== {date_input} の打刻記録 ===")
                print(f"{'時刻':<20} {'講師名':<15} {'種別':<6}")
                print("-" * 45)
                
                for name, record_type, timestamp in results:
                    action = "出勤" if record_type == "IN" else "退勤"
                    # 時刻部分のみを抽出
                    time_part = timestamp.split()[1] if ' ' in timestamp else timestamp
                    padded_name = self.pad_string(name, 15)
                    print(f"{time_part:<20} {padded_name} {action}")
                
                print(f"\n合計記録数: {len(results)}件")
            else:
                print(f"{date_input} の打刻記録はありません。")
                
        except Exception as e:
            print(f"指定日付の記録取得エラー: {e}")
    
    def show_date_summary(self):
        """特定の日付の出退勤状況サマリー"""
        print("\n=== 特定の日付の出退勤状況サマリー ===")
        
        # 日付入力
        while True:
            date_input = input("日付を入力してください (YYYY-MM-DD形式、例: 2025-01-15): ").strip()
            
            if not date_input:
                print("日付の入力をキャンセルしました。")
                return
            
            # 日付形式の検証
            try:
                # 日付の妥当性をチェック
                datetime.strptime(date_input, "%Y-%m-%d")
                break
            except ValueError:
                print("無効な日付形式です。YYYY-MM-DD形式で入力してください。")
                continue
        
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT instructor_name, record_type, timestamp
                FROM time_records
                WHERE DATE(timestamp) = ?
                ORDER BY instructor_name, timestamp
            ''', (date_input,))
            
            results = cursor.fetchall()
            conn.close()
            
            if results:
                # 講師ごとにグループ化
                instructor_records = {}
                for name, record_type, timestamp in results:
                    if name not in instructor_records:
                        instructor_records[name] = []
                    instructor_records[name].append((record_type, timestamp))
                
                print(f"\n=== {date_input} の出退勤状況サマリー ===")
                # ヘッダーの表示幅を調整
                header_name = self.pad_string("講師名", 15)
                header_status = self.pad_string("最終状態", 10)
                header_time = self.pad_string("最終打刻時刻", 20)
                print(f"{header_name} {header_status} {header_time} その日の記録")
                print("-" * 75)
                
                for name, records in instructor_records.items():
                    last_record = records[-1]
                    status = "出勤で終了" if last_record[0] == "IN" else "退勤で終了"
                    last_time = last_record[1]
                    
                    # その日の記録を文字列化
                    record_str = ""
                    for record_type, timestamp in records:
                        time_only = timestamp.split()[1][:5] if ' ' in timestamp else timestamp[:5]  # HH:MM形式
                        action = "出" if record_type == "IN" else "退"
                        record_str += f"{action}:{time_only} "
                    
                    # 各列の表示幅を調整
                    padded_name = self.pad_string(name, 15)
                    padded_status = self.pad_string(status, 10)
                    padded_time = self.pad_string(last_time, 20)
                    
                    print(f"{padded_name} {padded_status} {padded_time} {record_str}")
                
                print(f"\n{date_input} に打刻した講師数: {len(instructor_records)}人")
            else:
                print(f"{date_input} の記録はありません。")
                
        except Exception as e:
            print(f"指定日付のサマリー取得エラー: {e}")
    
    def get_display_width(self, text):
        """文字列の表示幅を計算（日本語文字は2文字分、半角文字は1文字分）"""
        width = 0
        for char in text:
            # 日本語文字（ひらがな、カタカナ、漢字、全角記号など）は2文字分
            if ord(char) > 127:  # ASCII文字以外
                width += 2
            else:
                width += 1
        return width
    
    def pad_string(self, text, target_width):
        """文字列を指定した表示幅になるように半角スペースでパディング"""
        current_width = self.get_display_width(text)
        if current_width < target_width:
            return text + " " * (target_width - current_width)
        return text
    
    def show_today_summary(self):
        """今日の出退勤状況サマリー"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT instructor_name, record_type, timestamp
                FROM time_records
                WHERE DATE(timestamp) = DATE('now', 'localtime')
                ORDER BY instructor_name, timestamp
            ''')
            
            results = cursor.fetchall()
            conn.close()
            
            if results:
                # 講師ごとにグループ化
                instructor_records = {}
                for name, record_type, timestamp in results:
                    if name not in instructor_records:
                        instructor_records[name] = []
                    instructor_records[name].append((record_type, timestamp))
                
                print("\n=== 今日の出退勤状況サマリー ===")
                # ヘッダーの表示幅を調整
                header_name = self.pad_string("講師名", 15)
                header_status = self.pad_string("状態", 8)
                header_time = self.pad_string("最終打刻時刻", 20)
                print(f"{header_name} {header_status} {header_time} 今日の記録")
                print("-" * 70)
                
                for name, records in instructor_records.items():
                    last_record = records[-1]
                    status = "出勤中" if last_record[0] == "IN" else "退勤済"
                    last_time = last_record[1]
                    
                    # 今日の記録を文字列化
                    record_str = ""
                    for record_type, timestamp in records:
                        time_only = timestamp.split()[1][:5]  # HH:MM形式
                        action = "出" if record_type == "IN" else "退"
                        record_str += f"{action}:{time_only} "
                    
                    # 各列の表示幅を調整
                    padded_name = self.pad_string(name, 15)
                    padded_status = self.pad_string(status, 8)
                    padded_time = self.pad_string(last_time, 20)
                    
                    print(f"{padded_name} {padded_status} {padded_time} {record_str}")
                
                print(f"\n今日打刻した講師数: {len(instructor_records)}人")
            else:
                print("今日の記録はありません。")
                
        except Exception as e:
            print(f"今日のサマリー取得エラー: {e}")
    
    def export_today_records_to_csv(self):
        """今日の打刻記録をCSVファイルにエクスポート"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 今日の日付を取得
            today = datetime.now(JST).strftime("%Y-%m-%d")
            
            cursor.execute('''
                SELECT instructor_name, record_type, timestamp, card_uid
                FROM time_records
                WHERE DATE(timestamp) = DATE('now', 'localtime')
                ORDER BY timestamp
            ''')
            
            results = cursor.fetchall()
            conn.close()
            
            if not results:
                print("今日の打刻記録はありません。エクスポートできません。")
                return
            
            # CSVファイル名を生成（重複回避のため連番を付ける）
            csv_filename = self.generate_unique_csv_filename(today)
            
            # CSVファイルに書き込み
            with open(csv_filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.writer(csvfile)
                
                # ヘッダー行を書き込み
                writer.writerow(['講師名', '打刻種別', '打刻日時', 'カードUID'])
                
                # データ行を書き込み
                for name, record_type, timestamp, card_uid in results:
                    # 打刻種別を日本語に変換
                    record_type_jp = "出勤" if record_type == "IN" else "退勤"
                    writer.writerow([name, record_type_jp, timestamp, card_uid])
            
            print(f"\n=== CSVエクスポート完了 ===")
            print(f"ファイル名: {csv_filename}")
            print(f"エクスポート件数: {len(results)}件")
            print(f"対象日: {today}")
            
            # エクスポートしたデータの概要を表示
            print(f"\n=== エクスポート内容概要 ===")
            instructor_count = len(set(name for name, _, _, _ in results))
            in_count = sum(1 for _, record_type, _, _ in results if record_type == "IN")
            out_count = sum(1 for _, record_type, _, _ in results if record_type == "OUT")
            
            print(f"打刻した講師数: {instructor_count}人")
            print(f"出勤記録: {in_count}件")
            print(f"退勤記録: {out_count}件")
            print(f"合計記録数: {len(results)}件")
            
        except Exception as e:
            print(f"CSVエクスポートエラー: {e}")
    
    def generate_unique_csv_filename(self, date_str):
        """重複しないCSVファイル名を生成"""
        # logディレクトリを作成（存在しない場合）
        log_dir = "log"
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        base_filename = f"attendance_records_{date_str}"
        csv_filename = os.path.join(log_dir, f"{base_filename}.csv")
        
        # ファイルが存在しない場合はそのまま返す
        if not os.path.exists(csv_filename):
            return csv_filename
        
        # ファイルが存在する場合は連番を付ける
        counter = 2
        while True:
            csv_filename = os.path.join(log_dir, f"{base_filename}_{counter}.csv")
            if not os.path.exists(csv_filename):
                return csv_filename
            counter += 1
    
    def export_date_records_to_csv(self):
        """特定の日付の打刻記録をCSVファイルにエクスポート"""
        print("\n=== 特定の日付の打刻記録CSVエクスポート ===")
        
        # 日付入力
        while True:
            date_input = input("日付を入力してください (YYYY-MM-DD形式、例: 2025-01-15): ").strip()
            
            if not date_input:
                print("日付の入力をキャンセルしました。")
                return
            
            # 日付形式の検証
            try:
                # 日付の妥当性をチェック
                datetime.strptime(date_input, "%Y-%m-%d")
                break
            except ValueError:
                print("無効な日付形式です。YYYY-MM-DD形式で入力してください。")
                continue
        
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT instructor_name, record_type, timestamp, card_uid
                FROM time_records
                WHERE DATE(timestamp) = ?
                ORDER BY timestamp
            ''', (date_input,))
            
            results = cursor.fetchall()
            conn.close()
            
            if not results:
                print(f"{date_input} の打刻記録はありません。エクスポートできません。")
                return
            
            # CSVファイル名を生成（重複回避のため連番を付ける）
            csv_filename = self.generate_unique_csv_filename(date_input)
            
            # CSVファイルに書き込み
            with open(csv_filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.writer(csvfile)
                
                # ヘッダー行を書き込み
                writer.writerow(['講師名', '打刻種別', '打刻日時', 'カードUID'])
                
                # データ行を書き込み
                for name, record_type, timestamp, card_uid in results:
                    # 打刻種別を日本語に変換
                    record_type_jp = "出勤" if record_type == "IN" else "退勤"
                    writer.writerow([name, record_type_jp, timestamp, card_uid])
            
            print(f"\n=== CSVエクスポート完了 ===")
            print(f"ファイル名: {csv_filename}")
            print(f"エクスポート件数: {len(results)}件")
            print(f"対象日: {date_input}")
            
            # エクスポートしたデータの概要を表示
            print(f"\n=== エクスポート内容概要 ===")
            instructor_count = len(set(name for name, _, _, _ in results))
            in_count = sum(1 for _, record_type, _, _ in results if record_type == "IN")
            out_count = sum(1 for _, record_type, _, _ in results if record_type == "OUT")
            
            print(f"打刻した講師数: {instructor_count}人")
            print(f"出勤記録: {in_count}件")
            print(f"退勤記録: {out_count}件")
            print(f"合計記録数: {len(results)}件")
            
        except Exception as e:
            print(f"CSVエクスポートエラー: {e}")
    
    def register_instructor_interactive(self):
        """インタラクティブな講師登録機能"""
        print("\n=== 講師登録 ===")
        print("カードをリーダーにかざしてください...")
        print("（キャンセルする場合は Ctrl+C を押してください）")
        
        try:
            # カード読み取り待機
            card_detected = False
            while not card_detected:
                if self.is_card_present():
                    if self.connect_to_card():
                        uid = self.get_card_uid()
                        if uid:
                            print(f"\nカードを検出しました！")
                            print(f"カードUID: {uid}")
                            
                            # 既存講師チェック
                            existing_instructor = self.get_instructor_by_uid(uid)
                            if existing_instructor:
                                print(f"このカードは既に登録されています: {existing_instructor}")
                                print("登録をキャンセルします。")
                                self.disconnect()
                                return
                            
                            # 講師情報入力
                            print("\n講師情報を入力してください:")
                            
                            # 講師番号の入力（オプション）
                            next_id = self.get_next_instructor_id()
                            instructor_id_input = input(f"講師番号 (デフォルト: {next_id}): ").strip()
                            
                            if instructor_id_input:
                                try:
                                    instructor_id = int(instructor_id_input)
                                    # 重複チェック
                                    if self.is_instructor_id_exists(instructor_id):
                                        print(f"講師番号 {instructor_id} は既に使用されています。")
                                        print(f"自動的に {next_id} を使用します。")
                                        instructor_id = next_id
                                except ValueError:
                                    print("無効な講師番号です。自動採番を使用します。")
                                    instructor_id = next_id
                            else:
                                instructor_id = next_id
                            
                            # 講師名の入力
                            name = input("講師名: ").strip()
                            if not name:
                                print("講師名が入力されていません。登録をキャンセルします。")
                                self.disconnect()
                                return
                            
                            # 確認
                            print(f"\n=== 登録内容確認 ===")
                            print(f"講師番号: {instructor_id}")
                            print(f"講師名: {name}")
                            print(f"カードUID: {uid}")
                            
                            confirm = input("\nこの内容で登録しますか？ (y/n): ").strip().lower()
                            if confirm == 'y' or confirm == 'yes':
                                # 手動で講師番号を指定する場合の登録
                                if self.add_instructor_with_id(instructor_id, uid, name):
                                    print("講師登録が完了しました！")
                                else:
                                    print("講師登録に失敗しました。")
                            else:
                                print("登録をキャンセルしました。")
                            
                            card_detected = True
                        self.disconnect()
                else:
                    time.sleep(0.1)  # 短い間隔でチェック
                    
        except KeyboardInterrupt:
            print("\n講師登録をキャンセルしました。")
            self.disconnect()
    
    def is_instructor_id_exists(self, instructor_id):
        """指定された講師番号が既に存在するかチェック"""
        try:
            if not os.path.exists(self.instructors_csv):
                return False
            
            with open(self.instructors_csv, 'r', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    try:
                        if int(row['instructor_id']) == instructor_id:
                            return True
                    except (ValueError, KeyError):
                        continue
            return False
        except Exception as e:
            print(f"講師番号チェックエラー: {e}")
            return False
    
    def add_instructor_with_id(self, instructor_id, card_uid, name):
        """指定された講師番号で講師を追加"""
        try:
            # 既存データをチェック
            instructors = self.load_instructors()
            if card_uid in instructors:
                print(f"エラー: カードUID {card_uid} は既に登録されています。")
                return False
            
            # 講師番号の重複チェック
            if self.is_instructor_id_exists(instructor_id):
                print(f"エラー: 講師番号 {instructor_id} は既に使用されています。")
                return False
            
            # CSVファイルに追加
            with open(self.instructors_csv, 'a', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                created_at = datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")
                writer.writerow([instructor_id, card_uid, name, created_at])
            
            print(f"講師を登録しました: {name} (講師番号: {instructor_id}, UID: {card_uid})")
            return True
            
        except Exception as e:
            print(f"講師登録エラー: {e}")
            return False
    
    def edit_instructors_csv(self):
        """講師マスタCSVファイルを編集用に開く"""
        try:
            if os.path.exists(self.instructors_csv):
                print(f"講師マスタファイル '{self.instructors_csv}' を確認してください。")
                print("このファイルを直接編集して講師情報を管理できます。")
                
                # ファイルの内容を表示
                with open(self.instructors_csv, 'r', encoding='utf-8') as csvfile:
                    content = csvfile.read()
                    print(f"\n現在の内容:\n{content}")
            else:
                print("講師マスタファイルが存在しません。")
        except Exception as e:
            print(f"ファイル確認エラー: {e}")
    
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
    system = AttendanceSystemCSV()
    
    if not system.initialize_reader():
        return
    
    while True:
        sound_status = "有効" if system.sound_enabled else "無効"
        print("\n=== 出退勤管理システム（CSV版） ===")
        print("1. 出退勤監視開始")
        print("2. 講師登録")
        print("3. 講師一覧表示")
        print("4. 今日の打刻記録表示")
        print("5. 今日の出退勤状況サマリー")
        print("6. 特定の日付の打刻記録表示")
        print("7. 特定の日付の出退勤状況サマリー")
        print("8. 講師マスタCSVファイル確認")
        print(f"9. 音声設定 (現在: {sound_status})")
        print("10. 今日の打刻記録をCSVエクスポート")
        print("11. 特定の日付の打刻記録をCSVエクスポート")
        print("12. 終了")
        
        choice = input("選択してください (1-12): ").strip()
        
        if choice == "1":
            try:
                system.monitor_attendance()
            except KeyboardInterrupt:
                print("\n監視を終了しました。")
                
        elif choice == "2":
            system.register_instructor_interactive()
                
        elif choice == "3":
            system.list_instructors()
            
        elif choice == "4":
            system.show_today_records()
            
        elif choice == "5":
            system.show_today_summary()
            
        elif choice == "6":
            system.show_date_records()
            
        elif choice == "7":
            system.show_date_summary()
            
        elif choice == "8":
            system.edit_instructors_csv()
            
        elif choice == "9":
            system.toggle_sound()
            
        elif choice == "10":
            system.export_today_records_to_csv()
            
        elif choice == "11":
            system.export_date_records_to_csv()
            
        elif choice == "12":
            print("システムを終了します。")
            break
            
        else:
            print("無効な選択です。")

if __name__ == "__main__":
    main()
