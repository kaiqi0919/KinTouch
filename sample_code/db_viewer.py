# attendance.dbの中身を人間が読みやすく表示するスクリプト

import sqlite3
from datetime import datetime

class DatabaseViewer:
    def __init__(self, db_path="attendance.db"):
        self.db_path = db_path
    
    def show_instructors(self):
        """講師マスタの内容を表示"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute("SELECT id, card_uid, name, created_at FROM instructors ORDER BY id")
            results = cursor.fetchall()
            conn.close()
            
            print("=" * 80)
            print("講師マスタテーブル (instructors)")
            print("=" * 80)
            
            if results:
                print(f"{'ID':<5} {'カードUID':<20} {'講師名':<15} {'登録日時':<20}")
                print("-" * 80)
                for row in results:
                    print(f"{row[0]:<5} {row[1]:<20} {row[2]:<15} {row[3]:<20}")
                print(f"\n合計: {len(results)}件")
            else:
                print("データがありません。")
            print()
            
        except Exception as e:
            print(f"講師マスタ取得エラー: {e}")
    
    def show_time_records(self, limit=None):
        """打刻記録の内容を表示"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            query = '''
                SELECT tr.id, i.name, tr.card_uid, tr.record_type, tr.timestamp
                FROM time_records tr
                LEFT JOIN instructors i ON tr.instructor_id = i.id
                ORDER BY tr.timestamp DESC
            '''
            
            if limit:
                query += f" LIMIT {limit}"
            
            cursor.execute(query)
            results = cursor.fetchall()
            conn.close()
            
            print("=" * 90)
            print(f"打刻記録テーブル (time_records){' - 最新' + str(limit) + '件' if limit else ''}")
            print("=" * 90)
            
            if results:
                print(f"{'ID':<5} {'講師名':<15} {'カードUID':<20} {'種別':<6} {'打刻日時':<20}")
                print("-" * 90)
                for row in results:
                    record_type_jp = "出勤" if row[3] == "IN" else "退勤"
                    instructor_name = row[1] if row[1] else "未登録"
                    print(f"{row[0]:<5} {instructor_name:<15} {row[2]:<20} {record_type_jp:<6} {row[4]:<20}")
                print(f"\n合計: {len(results)}件")
            else:
                print("データがありません。")
            print()
            
        except Exception as e:
            print(f"打刻記録取得エラー: {e}")
    
    def show_today_summary(self):
        """今日の出退勤状況サマリー"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 今日の記録を取得
            cursor.execute('''
                SELECT i.name, tr.record_type, tr.timestamp
                FROM time_records tr
                JOIN instructors i ON tr.instructor_id = i.id
                WHERE DATE(tr.timestamp) = DATE('now', 'localtime')
                ORDER BY i.name, tr.timestamp
            ''')
            
            results = cursor.fetchall()
            conn.close()
            
            print("=" * 70)
            print("今日の出退勤状況サマリー")
            print("=" * 70)
            
            if results:
                # 講師ごとにグループ化
                instructor_records = {}
                for name, record_type, timestamp in results:
                    if name not in instructor_records:
                        instructor_records[name] = []
                    instructor_records[name].append((record_type, timestamp))
                
                print(f"{'講師名':<15} {'状態':<8} {'最終打刻時刻':<20} {'今日の記録'}")
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
                    
                    print(f"{name:<15} {status:<8} {last_time:<20} {record_str}")
                
                print(f"\n今日打刻した講師数: {len(instructor_records)}人")
            else:
                print("今日の記録はありません。")
            print()
            
        except Exception as e:
            print(f"今日のサマリー取得エラー: {e}")
    
    def show_instructor_detail(self, instructor_name):
        """特定講師の詳細記録を表示"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 講師情報を取得
            cursor.execute("SELECT id, card_uid, created_at FROM instructors WHERE name = ?", (instructor_name,))
            instructor = cursor.fetchone()
            
            if not instructor:
                print(f"講師 '{instructor_name}' が見つかりません。")
                return
            
            # 打刻記録を取得
            cursor.execute('''
                SELECT record_type, timestamp
                FROM time_records
                WHERE instructor_id = ?
                ORDER BY timestamp DESC
                LIMIT 20
            ''', (instructor[0],))
            
            records = cursor.fetchall()
            conn.close()
            
            print("=" * 60)
            print(f"講師詳細: {instructor_name}")
            print("=" * 60)
            print(f"講師ID: {instructor[0]}")
            print(f"カードUID: {instructor[1]}")
            print(f"登録日: {instructor[2]}")
            print()
            
            if records:
                print("最新20件の打刻記録:")
                print(f"{'種別':<6} {'打刻日時':<20}")
                print("-" * 30)
                for record_type, timestamp in records:
                    action = "出勤" if record_type == "IN" else "退勤"
                    print(f"{action:<6} {timestamp:<20}")
            else:
                print("打刻記録がありません。")
            print()
            
        except Exception as e:
            print(f"講師詳細取得エラー: {e}")
    
    def show_statistics(self):
        """統計情報を表示"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 基本統計
            cursor.execute("SELECT COUNT(*) FROM instructors")
            instructor_count = cursor.fetchone()[0]
            
            cursor.execute("SELECT COUNT(*) FROM time_records")
            total_records = cursor.fetchone()[0]
            
            cursor.execute("SELECT COUNT(*) FROM time_records WHERE DATE(timestamp) = DATE('now', 'localtime')")
            today_records = cursor.fetchone()[0]
            
            cursor.execute("SELECT COUNT(DISTINCT instructor_id) FROM time_records WHERE DATE(timestamp) = DATE('now', 'localtime')")
            today_instructors = cursor.fetchone()[0]
            
            conn.close()
            
            print("=" * 50)
            print("統計情報")
            print("=" * 50)
            print(f"登録講師数: {instructor_count}人")
            print(f"総打刻記録数: {total_records}件")
            print(f"今日の打刻記録数: {today_records}件")
            print(f"今日打刻した講師数: {today_instructors}人")
            print()
            
        except Exception as e:
            print(f"統計情報取得エラー: {e}")

def main():
    """メイン関数"""
    viewer = DatabaseViewer()
    
    while True:
        print("=" * 50)
        print("出退勤データベース ビューアー")
        print("=" * 50)
        print("1. 講師マスタ表示")
        print("2. 打刻記録表示（全件）")
        print("3. 打刻記録表示（最新50件）")
        print("4. 今日の出退勤状況")
        print("5. 講師詳細表示")
        print("6. 統計情報")
        print("7. 全データ表示")
        print("8. 終了")
        
        choice = input("\n選択してください (1-8): ").strip()
        
        if choice == "1":
            viewer.show_instructors()
            
        elif choice == "2":
            viewer.show_time_records()
            
        elif choice == "3":
            viewer.show_time_records(50)
            
        elif choice == "4":
            viewer.show_today_summary()
            
        elif choice == "5":
            name = input("講師名を入力してください: ").strip()
            if name:
                viewer.show_instructor_detail(name)
            
        elif choice == "6":
            viewer.show_statistics()
            
        elif choice == "7":
            viewer.show_statistics()
            viewer.show_instructors()
            viewer.show_time_records()
            
        elif choice == "8":
            print("ビューアーを終了します。")
            break
            
        else:
            print("無効な選択です。")
        
        input("Enterキーを押して続行...")

if __name__ == "__main__":
    main()
