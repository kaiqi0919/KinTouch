#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
講師情報をCSVからデータベースに移行するスクリプト
"""

import sqlite3
import csv
import os
import shutil
from datetime import datetime

def migrate_instructors():
    """講師情報をCSVからDBに移行"""
    
    db_path = "data/attendance.db"
    csv_path = "data/instructors.csv"
    
    if not os.path.exists(csv_path):
        print(f"エラー: {csv_path} が見つかりません")
        return False
    
    try:
        # データベース接続
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # instructorsテーブルの存在確認
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='instructors'")
        table_exists = cursor.fetchone()
        
        if table_exists:
            print("警告: instructorsテーブルは既に存在します")
            print("既存のテーブルを削除して再作成します...")
            
            # 既存テーブルを削除
            cursor.execute("DROP TABLE instructors")
            print("既存のinstructorsテーブルを削除しました")
        
        # instructorsテーブルを作成
        cursor.execute('''
            CREATE TABLE instructors (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                instructor_id INTEGER UNIQUE NOT NULL,
                card_uid TEXT UNIQUE NOT NULL,
                name TEXT NOT NULL,
                created_at TIMESTAMP NOT NULL
            )
        ''')
        print("instructorsテーブルを作成しました")
        
        # CSVからデータを読み込み
        instructors = []
        with open(csv_path, 'r', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                instructors.append({
                    'instructor_id': int(row['instructor_id']),
                    'card_uid': row['card_uid'],
                    'name': row['name'].strip(),  # 改行コードも除去
                    'created_at': row['created_at']
                })
        
        print(f"\nCSVから{len(instructors)}名の講師データを読み込みました")
        
        # データベースに挿入
        for instructor in instructors:
            cursor.execute('''
                INSERT INTO instructors (instructor_id, card_uid, name, created_at)
                VALUES (?, ?, ?, ?)
            ''', (
                instructor['instructor_id'],
                instructor['card_uid'],
                instructor['name'],
                instructor['created_at']
            ))
            print(f"OK [{instructor['instructor_id']}] {instructor['name']}")
        
        conn.commit()
        print(f"\n{len(instructors)}名の講師データをデータベースに登録しました")
        
        # CSVファイルをバックアップ
        backup_path = csv_path + ".backup_" + datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy2(csv_path, backup_path)
        print(f"\n元のCSVファイルを {backup_path} にバックアップしました")
        
        # 確認: データベースから読み込み
        cursor.execute("SELECT COUNT(*) FROM instructors")
        count = cursor.fetchone()[0]
        print(f"\nデータベース確認: {count}名の講師が登録されています")
        
        conn.close()
        
        print("\n" + "="*60)
        print("移行が正常に完了しました！")
        print("="*60)
        print("\n次のステップ:")
        print("1. 出退勤確認システム.py を更新してDBから読み込むように変更")
        print("2. 動作確認後、data/instructors.csv を削除または移動")
        
        return True
        
    except Exception as e:
        print(f"\nエラーが発生しました: {e}")
        if 'conn' in locals():
            conn.rollback()
            conn.close()
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("講師情報 CSV→DB 移行ツール")
    print("=" * 60)
    print()
    
    migrate_instructors()