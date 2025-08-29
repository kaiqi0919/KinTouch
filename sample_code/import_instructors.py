# 講師マスタをCSVファイルからインポートするスクリプト

import csv
import sqlite3
import os

class InstructorImporter:
    def __init__(self, db_path="attendance.db"):
        self.db_path = db_path
    
    def create_sample_csv(self, filename="instructors_sample.csv"):
        """サンプルCSVファイルを作成"""
        sample_data = [
            ["card_uid", "name"],
            ["04 12 34 56 78 9A BC", "田中太郎"],
            ["04 AB CD EF 12 34 56", "佐藤花子"],
            ["04 98 76 54 32 10 FE", "鈴木一郎"],
            ["04 11 22 33 44 55 66", "山田美咲"]
        ]
        
        try:
            with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerows(sample_data)
            
            print(f"サンプルCSVファイル '{filename}' を作成しました。")
            print("このファイルを編集して実際の講師データを入力してください。")
            return True
            
        except Exception as e:
            print(f"サンプルCSVファイル作成エラー: {e}")
            return False
    
    def validate_csv(self, filename):
        """CSVファイルの形式をチェック"""
        if not os.path.exists(filename):
            print(f"ファイル '{filename}' が見つかりません。")
            return False
        
        try:
            with open(filename, 'r', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                header = next(reader)
                
                # ヘッダーチェック
                if len(header) < 2:
                    print("CSVファイルには最低2列（card_uid, name）が必要です。")
                    return False
                
                if header[0].lower() != 'card_uid' or header[1].lower() != 'name':
                    print("CSVファイルのヘッダーは 'card_uid,name' である必要があります。")
                    return False
                
                # データ行をチェック
                row_count = 0
                for row_num, row in enumerate(reader, start=2):
                    if len(row) < 2:
                        print(f"行 {row_num}: データが不足しています。")
                        return False
                    
                    if not row[0].strip() or not row[1].strip():
                        print(f"行 {row_num}: card_uidまたはnameが空です。")
                        return False
                    
                    row_count += 1
                
                if row_count == 0:
                    print("CSVファイルにデータ行がありません。")
                    return False
                
                print(f"CSVファイルの検証完了: {row_count}件のデータが見つかりました。")
                return True
                
        except Exception as e:
            print(f"CSVファイル検証エラー: {e}")
            return False
    
    def import_from_csv(self, filename, skip_duplicates=True):
        """CSVファイルから講師データをインポート"""
        if not self.validate_csv(filename):
            return False
        
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            imported_count = 0
            skipped_count = 0
            error_count = 0
            
            with open(filename, 'r', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                next(reader)  # ヘッダーをスキップ
                
                for row_num, row in enumerate(reader, start=2):
                    card_uid = row[0].strip()
                    name = row[1].strip()
                    
                    try:
                        if skip_duplicates:
                            # 重複チェック
                            cursor.execute("SELECT id FROM instructors WHERE card_uid = ?", (card_uid,))
                            if cursor.fetchone():
                                print(f"行 {row_num}: カードUID '{card_uid}' は既に登録済みです。スキップします。")
                                skipped_count += 1
                                continue
                        
                        # データ挿入
                        cursor.execute(
                            "INSERT INTO instructors (card_uid, name) VALUES (?, ?)",
                            (card_uid, name)
                        )
                        
                        print(f"行 {row_num}: {name} (UID: {card_uid}) を登録しました。")
                        imported_count += 1
                        
                    except sqlite3.IntegrityError as e:
                        print(f"行 {row_num}: 重複エラー - {e}")
                        error_count += 1
                    except Exception as e:
                        print(f"行 {row_num}: インポートエラー - {e}")
                        error_count += 1
            
            conn.commit()
            conn.close()
            
            print(f"\nインポート完了:")
            print(f"  成功: {imported_count}件")
            print(f"  スキップ: {skipped_count}件")
            print(f"  エラー: {error_count}件")
            
            return imported_count > 0
            
        except Exception as e:
            print(f"インポート処理エラー: {e}")
            return False
    
    def export_to_csv(self, filename="instructors_export.csv"):
        """現在の講師データをCSVファイルにエクスポート"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute("SELECT card_uid, name, created_at FROM instructors ORDER BY name")
            results = cursor.fetchall()
            conn.close()
            
            with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["card_uid", "name", "created_at"])
                writer.writerows(results)
            
            print(f"講師データを '{filename}' にエクスポートしました。({len(results)}件)")
            return True
            
        except Exception as e:
            print(f"エクスポートエラー: {e}")
            return False
    
    def show_current_instructors(self):
        """現在登録されている講師を表示"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute("SELECT card_uid, name, created_at FROM instructors ORDER BY name")
            results = cursor.fetchall()
            conn.close()
            
            if results:
                print(f"\n現在登録されている講師 ({len(results)}人):")
                print("-" * 70)
                print(f"{'カードUID':<20} {'講師名':<15} {'登録日時':<20}")
                print("-" * 70)
                for uid, name, created_at in results:
                    print(f"{uid:<20} {name:<15} {created_at:<20}")
            else:
                print("登録されている講師はいません。")
            print()
            
        except Exception as e:
            print(f"講師一覧取得エラー: {e}")

def main():
    """メイン関数"""
    importer = InstructorImporter()
    
    while True:
        print("=" * 60)
        print("講師マスタ インポート/エクスポート ツール")
        print("=" * 60)
        print("1. サンプルCSVファイル作成")
        print("2. CSVファイルから講師をインポート")
        print("3. 講師データをCSVファイルにエクスポート")
        print("4. 現在の講師一覧表示")
        print("5. CSVファイル形式の説明")
        print("6. 終了")
        
        choice = input("\n選択してください (1-6): ").strip()
        
        if choice == "1":
            filename = input("作成するCSVファイル名 (デフォルト: instructors_sample.csv): ").strip()
            if not filename:
                filename = "instructors_sample.csv"
            importer.create_sample_csv(filename)
            
        elif choice == "2":
            filename = input("インポートするCSVファイル名: ").strip()
            if filename:
                skip_duplicates = input("重複するカードUIDをスキップしますか？ (y/n, デフォルト: y): ").strip().lower()
                skip_duplicates = skip_duplicates != 'n'
                importer.import_from_csv(filename, skip_duplicates)
            
        elif choice == "3":
            filename = input("エクスポート先CSVファイル名 (デフォルト: instructors_export.csv): ").strip()
            if not filename:
                filename = "instructors_export.csv"
            importer.export_to_csv(filename)
            
        elif choice == "4":
            importer.show_current_instructors()
            
        elif choice == "5":
            print("\n" + "=" * 60)
            print("CSVファイル形式の説明")
            print("=" * 60)
            print("CSVファイルは以下の形式で作成してください:")
            print()
            print("card_uid,name")
            print("04 12 34 56 78 9A BC,田中太郎")
            print("04 AB CD EF 12 34 56,佐藤花子")
            print("04 98 76 54 32 10 FE,鈴木一郎")
            print()
            print("注意事項:")
            print("- 1行目は必ずヘッダー行 (card_uid,name)")
            print("- card_uidは実際のカードから読み取った値を使用")
            print("- nameは講師の名前（日本語可）")
            print("- 文字エンコーディングはUTF-8")
            print("- 重複するcard_uidがある場合はスキップされます")
            print()
            
        elif choice == "6":
            print("ツールを終了します。")
            break
            
        else:
            print("無効な選択です。")
        
        input("Enterキーを押して続行...")

if __name__ == "__main__":
    main()
