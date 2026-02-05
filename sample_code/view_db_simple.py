#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""attendance.dbの中身を表示する簡易スクリプト"""

import sqlite3
import sys

def view_database(db_path="../data/attendance.db"):
    """データベースの内容を表示"""
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # 統計情報
        print("=" * 80)
        print("データベース統計情報")
        print("=" * 80)
        
        cursor.execute("SELECT COUNT(*) FROM time_records")
        time_records_count = cursor.fetchone()[0]
        print(f"授業打刻記録数: {time_records_count}件")
        
        cursor.execute("SELECT COUNT(*) FROM meeting_records")
        meeting_records_count = cursor.fetchone()[0]
        print(f"会議打刻記録数: {meeting_records_count}件")
        
        cursor.execute("SELECT COUNT(DISTINCT instructor_id) FROM time_records")
        time_instructors = cursor.fetchone()[0]
        print(f"授業に登録されている講師数: {time_instructors}人")
        
        cursor.execute("SELECT COUNT(DISTINCT instructor_id) FROM meeting_records")
        meeting_instructors = cursor.fetchone()[0]
        print(f"会議に登録されている講師数: {meeting_instructors}人")
        
        cursor.execute("SELECT COUNT(*) FROM time_records WHERE DATE(timestamp) = DATE('now', 'localtime')")
        today_time_records = cursor.fetchone()[0]
        print(f"今日の授業打刻記録数: {today_time_records}件")
        
        cursor.execute("SELECT COUNT(*) FROM meeting_records WHERE DATE(timestamp) = DATE('now', 'localtime')")
        today_meeting_records = cursor.fetchone()[0]
        print(f"今日の会議打刻記録数: {today_meeting_records}件\n")
        
        # 登録講師一覧（授業）
        print("=" * 90)
        print("登録講師一覧（授業 - time_records）")
        print("=" * 90)
        cursor.execute("""
            SELECT DISTINCT instructor_id, instructor_name, card_uid
            FROM time_records
            ORDER BY instructor_id
        """)
        time_instructors_list = cursor.fetchall()
        
        if time_instructors_list:
            print(f"{'講師ID':<10} {'講師名':<20} {'カードUID':<20}")
            print("-" * 90)
            for row in time_instructors_list:
                print(f"{row[0]:<10} {row[1]:<20} {row[2]:<20}")
            print(f"\n合計: {len(time_instructors_list)}人\n")
        else:
            print("データがありません。\n")
        
        # 登録講師一覧（会議）
        print("=" * 90)
        print("登録講師一覧（会議 - meeting_records）")
        print("=" * 90)
        cursor.execute("""
            SELECT DISTINCT instructor_id, instructor_name, card_uid
            FROM meeting_records
            ORDER BY instructor_id
        """)
        meeting_instructors_list = cursor.fetchall()
        
        if meeting_instructors_list:
            print(f"{'講師ID':<10} {'講師名':<20} {'カードUID':<20}")
            print("-" * 90)
            for row in meeting_instructors_list:
                print(f"{row[0]:<10} {row[1]:<20} {row[2]:<20}")
            print(f"\n合計: {len(meeting_instructors_list)}人\n")
        else:
            print("データがありません。\n")
        
        # 最新の授業打刻記録
        print("=" * 100)
        print("最新の授業打刻記録 (time_records) - 最新20件")
        print("=" * 100)
        
        cursor.execute("""
            SELECT id, instructor_name, card_uid, record_type, timestamp
            FROM time_records
            ORDER BY timestamp DESC
            LIMIT 20
        """)
        time_records = cursor.fetchall()
        
        if time_records:
            print(f"{'ID':<5} {'講師名':<20} {'カードUID':<20} {'種別':<6} {'打刻日時':<25}")
            print("-" * 100)
            for row in time_records:
                record_type_jp = "出勤" if row[3] == "IN" else "退勤"
                print(f"{row[0]:<5} {row[1]:<20} {row[2]:<20} {record_type_jp:<6} {row[4]:<25}")
            print(f"\n表示件数: {len(time_records)}件 / 総件数: {time_records_count}件\n")
        else:
            print("データがありません。\n")
        
        # 最新の会議打刻記録
        print("=" * 100)
        print("最新の会議打刻記録 (meeting_records) - 最新20件")
        print("=" * 100)
        
        cursor.execute("""
            SELECT id, instructor_name, card_uid, record_type, timestamp
            FROM meeting_records
            ORDER BY timestamp DESC
            LIMIT 20
        """)
        meeting_records = cursor.fetchall()
        
        if meeting_records:
            print(f"{'ID':<5} {'講師名':<20} {'カードUID':<20} {'種別':<6} {'打刻日時':<25}")
            print("-" * 100)
            for row in meeting_records:
                record_type_jp = "出勤" if row[3] == "IN" else "退勤"
                print(f"{row[0]:<5} {row[1]:<20} {row[2]:<20} {record_type_jp:<6} {row[4]:<25}")
            print(f"\n表示件数: {len(meeting_records)}件 / 総件数: {meeting_records_count}件\n")
        else:
            print("データがありません。\n")
        
        # 今日の授業出退勤状況
        print("=" * 90)
        print("今日の授業出退勤状況サマリー")
        print("=" * 90)
        
        cursor.execute("""
            SELECT instructor_name, record_type, timestamp
            FROM time_records
            WHERE DATE(timestamp) = DATE('now', 'localtime')
            ORDER BY instructor_name, timestamp
        """)
        today_time_data = cursor.fetchall()
        
        if today_time_data:
            instructor_records = {}
            for name, record_type, timestamp in today_time_data:
                if name not in instructor_records:
                    instructor_records[name] = []
                instructor_records[name].append((record_type, timestamp))
            
            print(f"{'講師名':<20} {'状態':<10} {'今日の記録'}")
            print("-" * 90)
            
            for name, records in instructor_records.items():
                last_record = records[-1]
                status = "出勤中" if last_record[0] == "IN" else "退勤済"
                
                record_str = ""
                for record_type, timestamp in records:
                    time_only = timestamp.split()[1][:5]
                    action = "出" if record_type == "IN" else "退"
                    record_str += f"{action}:{time_only} "
                
                print(f"{name:<20} {status:<10} {record_str}")
            
            print(f"\n今日打刻した講師数: {len(instructor_records)}人")
        else:
            print("今日の記録はありません。")
        
        print()
        
        # 今日の会議出退勤状況
        print("=" * 90)
        print("今日の会議出退勤状況サマリー")
        print("=" * 90)
        
        cursor.execute("""
            SELECT instructor_name, record_type, timestamp
            FROM meeting_records
            WHERE DATE(timestamp) = DATE('now', 'localtime')
            ORDER BY instructor_name, timestamp
        """)
        today_meeting_data = cursor.fetchall()
        
        if today_meeting_data:
            instructor_records = {}
            for name, record_type, timestamp in today_meeting_data:
                if name not in instructor_records:
                    instructor_records[name] = []
                instructor_records[name].append((record_type, timestamp))
            
            print(f"{'講師名':<20} {'状態':<10} {'今日の記録'}")
            print("-" * 90)
            
            for name, records in instructor_records.items():
                last_record = records[-1]
                status = "出勤中" if last_record[0] == "IN" else "退勤済"
                
                record_str = ""
                for record_type, timestamp in records:
                    time_only = timestamp.split()[1][:5]
                    action = "出" if record_type == "IN" else "退"
                    record_str += f"{action}:{time_only} "
                
                print(f"{name:<20} {status:<10} {record_str}")
            
            print(f"\n今日打刻した講師数: {len(instructor_records)}人")
        else:
            print("今日の記録はありません。")
        
        conn.close()
        print("\n" + "=" * 80)
        
    except sqlite3.Error as e:
        print(f"データベースエラー: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"エラー: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    import os
    # スクリプトのディレクトリを基準に絶対パスを構築
    script_dir = os.path.dirname(os.path.abspath(__file__))
    db_path = os.path.join(script_dir, "..", "data", "attendance.db")
    db_path = os.path.normpath(db_path)
    print(f"データベースパス: {db_path}")
    print(f"ファイル存在確認: {os.path.exists(db_path)}\n")
    view_database(db_path)
