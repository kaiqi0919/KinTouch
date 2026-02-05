# 出退勤管理システム - CSV出力モジュール

import csv
import os
import shutil
from datetime import datetime

class CSVExporter:
    """CSV出力管理クラス"""
    
    def __init__(self, db_manager):
        self.db_manager = db_manager
    
    def export_records_to_csv(self, date_str, table_name="time_records"):
        """日次CSVエクスポート"""
        try:
            # データ取得
            results = self.db_manager.get_date_records(date_str, table_name)
            
            if not results:
                return f"{date_str} の打刻記録はありません。"
            
            # テーブルタイプ
            table_type = "class" if table_name == "time_records" else "meeting"
            
            # CSVファイル名を生成
            csv_filename = self.generate_unique_csv_filename(date_str, table_type)
            
            # CSVファイルに書き込み
            with open(csv_filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.writer(csvfile)
                table_type_name = "授業用" if table_name == "time_records" else "会議用"
                writer.writerow([f'【{table_type_name}】講師名', '打刻種別', '打刻日時', 'カードUID'])
                
                for name, record_type, timestamp in results:
                    # card_uidは元のコードに合わせて空にする（results構造が異なるため）
                    record_type_jp = "出勤" if record_type == "IN" else "退勤"
                    writer.writerow([name, record_type_jp, timestamp, ''])
            
            # 統計情報
            instructor_count = len(set(name for name, _, _ in results))
            in_count = sum(1 for _, record_type, _ in results if record_type == "IN")
            out_count = sum(1 for _, record_type, _ in results if record_type == "OUT")
            
            table_type_name = "授業用" if table_name == "time_records" else "会議用"
            result = f"=== CSVエクスポート完了 ===\n\n"
            result += f"種別: {table_type_name}\n"
            result += f"ファイル名: {csv_filename}\n"
            result += f"対象日: {date_str}\n"
            result += f"エクスポート件数: {len(results)}件\n\n"
            result += f"=== エクスポート内容概要 ===\n"
            result += f"打刻した講師数: {instructor_count}人\n"
            result += f"出勤記録: {in_count}件\n"
            result += f"退勤記録: {out_count}件\n"
            result += f"合計記録数: {len(results)}件\n"
            
            return result
            
        except Exception as e:
            return f"CSVエクスポートエラー: {e}"
    
    def generate_unique_csv_filename(self, date_str, table_type="class"):
        """CSVファイル名を生成"""
        year_month = date_str[:7]
        
        daily_dir = "daily"
        month_dir = os.path.join(daily_dir, year_month)
        old_dir = os.path.join(month_dir, "old")
        
        if not os.path.exists(month_dir):
            os.makedirs(month_dir)
        if not os.path.exists(old_dir):
            os.makedirs(old_dir)
        
        type_prefix = "授業" if table_type == "class" else "会議"
        base_filename = f"【{type_prefix}】日次記録_{date_str}"
        csv_filename = os.path.join(month_dir, f"{base_filename}.csv")
        
        if os.path.exists(csv_filename):
            file_mtime = os.path.getmtime(csv_filename)
            file_datetime = datetime.fromtimestamp(file_mtime)
            timestamp_str = file_datetime.strftime("%H%M%S")
            
            old_filename = os.path.join(old_dir, f"{base_filename}_{timestamp_str}.csv")
            
            if os.path.exists(old_filename):
                counter = 2
                while True:
                    old_filename = os.path.join(old_dir, f"{base_filename}_{timestamp_str}_{counter}.csv")
                    if not os.path.exists(old_filename):
                        break
                    counter += 1
            
            shutil.move(csv_filename, old_filename)
        
        return csv_filename
    
    def export_instructor_daily_summary(self, month_str, instructor_id, instructor_name, output_dir, table_name="time_records"):
        """講師別日次集計CSVエクスポート"""
        try:
            import calendar
            
            # 月の日数を取得
            year, month = map(int, month_str.split('-'))
            _, last_day = calendar.monthrange(year, month)
            
            # データベースから打刻記録を取得
            records = self.db_manager.get_instructor_monthly_records(month_str, instructor_id, table_name)
            
            # 日付ごとにデータを整理
            daily_data = {}
            for date_str, time_str in records:
                if date_str not in daily_data:
                    daily_data[date_str] = []
                daily_data[date_str].append(time_str)
            
            # テーブルタイプに応じたファイル名プレフィックス
            table_type_prefix = "授業" if table_name == "time_records" else "会議"
            
            # CSVファイル名
            base_filename = f"【{table_type_prefix}】出退勤記録_{month_str}_{instructor_id}_{instructor_name}"
            csv_filename = os.path.join(output_dir, f"{base_filename}.csv")
            
            # 既存ファイルがある場合はoldフォルダに移動
            old_dir = os.path.join(output_dir, "old")
            if not os.path.exists(old_dir):
                os.makedirs(old_dir)
            
            if os.path.exists(csv_filename):
                file_mtime = os.path.getmtime(csv_filename)
                file_datetime = datetime.fromtimestamp(file_mtime)
                timestamp_str = file_datetime.strftime("%H%M%S")
                
                old_filename = os.path.join(old_dir, f"{base_filename}_{timestamp_str}.csv")
                
                if os.path.exists(old_filename):
                    counter = 2
                    while True:
                        old_filename = os.path.join(old_dir, f"{base_filename}_{timestamp_str}_{counter}.csv")
                        if not os.path.exists(old_filename):
                            break
                        counter += 1
                
                shutil.move(csv_filename, old_filename)
            
            # CSVファイルに書き込み
            with open(csv_filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(['講師ID', '講師名', '日付', '出社時刻', '退社時刻', '外出時刻', '復帰時刻', '備考'])
                
                # 月の各日について出力
                for day in range(1, last_day + 1):
                    date_obj = datetime(year, month, day)
                    date_str = date_obj.strftime('%Y-%m-%d')
                    
                    # その日の打刻時刻を取得
                    if date_str in daily_data:
                        times = sorted(daily_data[date_str])
                        start_time = times[0]   # 最も早い打刻
                        end_time = times[-1]    # 最も遅い打刻
                        
                        writer.writerow([
                            instructor_id,
                            instructor_name,
                            date_str,
                            start_time,
                            end_time,
                            '',  # 外出時刻(空欄)
                            '',  # 復帰時刻(空欄)
                            ''   # 備考(空欄)
                        ])
                    else:
                        # 打刻がない日は空欄
                        writer.writerow([
                            instructor_id,
                            instructor_name,
                            date_str,
                            '',  # 出社時刻(空欄)
                            '',  # 退社時刻(空欄)
                            '',  # 外出時刻(空欄)
                            '',  # 復帰時刻(空欄)
                            ''   # 備考(空欄)
                        ])
            
            return True
            
        except Exception as e:
            print(f"講師別日次集計エラー ({instructor_name}): {e}")
            return False