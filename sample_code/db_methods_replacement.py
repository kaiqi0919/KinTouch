#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
出退勤確認システム.py に追加するDB参照メソッド

CSVベースからDBベースに変更する際に使用
"""

# ===== 以下のメソッドを出退勤確認システム.pyに置き換える =====

def init_instructors_csv(self):
    """講師マスタDB初期化（CSVは不要）"""
    # instructorsテーブルはinit_databaseで作成されるため、
    # このメソッドは何もしない（後方互換性のため残す）
    pass

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