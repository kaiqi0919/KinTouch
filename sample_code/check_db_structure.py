#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""attendance.dbの構造を確認するスクリプト"""

import sqlite3
import sys
import os

def check_database():
    """データベースの構造を確認"""
    try:
        # スクリプトのディレクトリを基準に絶対パスを構築
        script_dir = os.path.dirname(os.path.abspath(__file__))
        db_path = os.path.join(script_dir, "..", "data", "attendance.db")
        db_path = os.path.normpath(db_path)
        
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        print("=" * 80)
        print(f"データベース: {db_path}")
        print("=" * 80)
        
        # テーブル一覧を取得
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
        tables = cursor.fetchall()
        
        print(f"\nテーブル数: {len(tables)}")
        print("-" * 80)
        
        if not tables:
            print("テーブルが存在しません。")
            conn.close()
            return
        
        # 各テーブルの情報を表示
        for (table_name,) in tables:
            print(f"\n【テーブル名: {table_name}】")
            
            # カラム情報を取得
            cursor.execute(f"PRAGMA table_info({table_name})")
            columns = cursor.fetchall()
            
            print(f"{'ID':<5} {'カラム名':<20} {'型':<15} {'NULL可':<10} {'デフォルト値':<15} {'主キー'}")
            print("-" * 80)
            for col in columns:
                cid, name, col_type, notnull, default_val, pk = col
                notnull_str = "NOT NULL" if notnull else "NULL"
                default_str = str(default_val) if default_val else ""
                pk_str = "PRIMARY KEY" if pk else ""
                print(f"{cid:<5} {name:<20} {col_type:<15} {notnull_str:<10} {default_str:<15} {pk_str}")
            
            # レコード数を取得
            cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
            count = cursor.fetchone()[0]
            print(f"\nレコード数: {count}件")
            
            # サンプルデータを表示（最大5件）
            if count > 0:
                cursor.execute(f"SELECT * FROM {table_name} LIMIT 5")
                samples = cursor.fetchall()
                
                print("\nサンプルデータ（最大5件）:")
                print("-" * 80)
                
                # カラム名のヘッダー
                col_names = [col[1] for col in columns]
                header = " | ".join([f"{name[:15]:<15}" for name in col_names])
                print(header)
                print("-" * 80)
                
                # データ
                for row in samples:
                    row_str = " | ".join([f"{str(val)[:15]:<15}" for val in row])
                    print(row_str)
        
        conn.close()
        print("\n" + "=" * 80)
        
    except sqlite3.Error as e:
        print(f"データベースエラー: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"エラー: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    check_database()