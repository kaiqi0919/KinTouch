# 出退勤管理システム - 月次集計モジュール

import csv
import os
import shutil
import calendar
from datetime import datetime

class MonthlyExporter:
    """月次集計管理クラス"""
    
    def __init__(self, db_manager, csv_exporter):
        self.db_manager = db_manager
        self.csv_exporter = csv_exporter
    
    def export_monthly_summary_to_csv(self, month_str, table_name="time_records", include_daily=True):
        """月次集計CSVエクスポート"""
        try:
            # 対象月の日付一覧を取得
            dates = self.db_manager.get_monthly_dates(month_str, table_name)
            
            if not dates:
                return f"{month_str} の打刻記録はありません。"
            
            # ステップ1: 各日の日次集計を実行（オプション）
            result = f"=== 月次集計処理開始 ===\n\n"
            result += f"対象月: {month_str}\n"
            result += f"打刻記録がある日数: {len(dates)}日\n\n"
            
            daily_export_count = 0
            if include_daily:
                result += f"--- 日次集計処理 ---\n"
                
                for date in dates:
                    daily_result = self.csv_exporter.export_records_to_csv(date, table_name)
                    if "エクスポート完了" in daily_result:
                        daily_export_count += 1
                        result += f"✓ {date}: 日次集計完了\n"
                    else:
                        result += f"✗ {date}: 日次集計失敗\n"
                
                result += f"\n日次集計完了: {daily_export_count}/{len(dates)}日\n\n"
            else:
                result += f"日次集計: スキップ\n\n"
            
            # ステップ2: 月次集計を実行
            result += f"--- 月次集計処理 ---\n"
            
            results = self.db_manager.get_monthly_summary_data(month_str, table_name)
            
            # 講師ごとに出勤日数をカウント
            instructor_attendance = {}
            for instructor_id, name, date in results:
                if name not in instructor_attendance:
                    instructor_attendance[name] = 0
                instructor_attendance[name] += 1
            
            # テーブルタイプ
            table_type_prefix = "授業" if table_name == "time_records" else "会議"
            
            # monthlyフォルダの準備
            monthly_dir = "monthly"
            month_subdir = os.path.join(monthly_dir, month_str)
            old_dir = os.path.join(month_subdir, "old")
            
            if not os.path.exists(month_subdir):
                os.makedirs(month_subdir)
            if not os.path.exists(old_dir):
                os.makedirs(old_dir)
            
            csv_filename = os.path.join(month_subdir, f"【{table_type_prefix}】出退勤記録_{month_str}.csv")
            
            # 既存ファイルがある場合はoldフォルダに移動
            self._move_to_old(csv_filename, old_dir, f"【{table_type_prefix}】出退勤記録_{month_str}")
            
            # すべての講師を取得してID順にソート
            all_instructors = self.db_manager.load_instructors_full()
            all_instructors_sorted = sorted(all_instructors, key=lambda x: int(x['instructor_id']))
            
            # まとめCSVファイルに書き込み
            year, month = map(int, month_str.split('-'))
            _, last_day = calendar.monthrange(year, month)
            
            with open(csv_filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(['講師ID', '講師名', '日付', '出社時刻', '退社時刻', '外出時刻', '復帰時刻', '備考'])
                
                # すべての講師について日次データを出力
                for instructor in all_instructors_sorted:
                    instructor_id = instructor['instructor_id']
                    instructor_name = instructor['name']
                    
                    # 打刻記録を取得
                    records_daily = self.db_manager.get_instructor_monthly_records(month_str, instructor_id, table_name)
                    
                    # 日付ごとにデータを整理
                    daily_data_per_instructor = {}
                    for date_str, time_str in records_daily:
                        if date_str not in daily_data_per_instructor:
                            daily_data_per_instructor[date_str] = []
                        daily_data_per_instructor[date_str].append(time_str)
                    
                    # 月の各日について出力
                    for day in range(1, last_day + 1):
                        date_obj = datetime(year, month, day)
                        date_str = date_obj.strftime('%Y-%m-%d')
                        
                        if date_str in daily_data_per_instructor:
                            times = sorted(daily_data_per_instructor[date_str])
                            start_time = times[0]
                            end_time = times[-1]
                            
                            writer.writerow([
                                instructor_id,
                                instructor_name,
                                date_str,
                                start_time,
                                end_time,
                                '', '', ''
                            ])
                        else:
                            writer.writerow([
                                instructor_id,
                                instructor_name,
                                date_str,
                                '', '', '', '', ''
                            ])
            
            # 統計情報
            total_instructors_registered = len(all_instructors_sorted)
            total_instructors_attended = len(instructor_attendance)
            total_days = sum(instructor_attendance.values())
            
            result += f"\n=== 月次集計エクスポート完了 ===\n\n"
            result += f"種別: {table_type_prefix}用\n"
            result += f"ファイル名: {csv_filename}\n"
            result += f"対象月: {month_str}\n\n"
            result += f"=== 集計結果 ===\n"
            result += f"登録講師数: {total_instructors_registered}人\n"
            result += f"出勤した講師数: {total_instructors_attended}人\n"
            result += f"総出勤日数: {total_days}日\n\n"
            result += f"=== 講師別出勤回数（ID順） ===\n"
            
            for instructor in all_instructors_sorted:
                name = instructor['name']
                instructor_id = instructor['instructor_id']
                count = instructor_attendance.get(name, 0)
                result += f"[{instructor_id}] {name}: {count}回\n"
            
            # ステップ3: 講師別日次集計を実行
            result += f"\n--- 講師別日次集計処理 ---\n"
            
            instructor_daily_count = 0
            for instructor in all_instructors_sorted:
                instructor_id = instructor['instructor_id']
                instructor_name = instructor['name']
                
                if self.csv_exporter.export_instructor_daily_summary(
                    month_str, instructor_id, instructor_name, month_subdir, table_name
                ):
                    instructor_daily_count += 1
                    result += f"✓ [{instructor_id}] {instructor_name}: 日次集計完了\n"
            
            result += f"\n講師別日次集計完了: {instructor_daily_count}/{total_instructors_registered}人\n"
            
            result += f"\n=== 処理完了 ===\n"
            if include_daily:
                result += f"日次集計ファイル: {daily_export_count}件作成\n"
            result += f"月次集計ファイル: 1件作成\n"
            result += f"講師別日次集計ファイル: {instructor_daily_count}件作成\n"
            
            return result
            
        except Exception as e:
            return f"月次集計エクスポートエラー: {e}"
    
    def export_combined_monthly_summary(self, month_str):
        """授業と会議を統合した月次集計CSVエクスポート"""
        try:
            # 授業記録・会議記録を取得
            class_results = self.db_manager.get_monthly_summary_data(month_str, "time_records")
            meeting_results = self.db_manager.get_monthly_summary_data(month_str, "meeting_records")
            
            # 講師ごとに集計
            instructor_summary = {}
            
            # 授業回数を集計
            for instructor_id, name, date in class_results:
                key = str(instructor_id)
                if key not in instructor_summary:
                    instructor_summary[key] = {'class_dates': set(), 'meeting_dates': set()}
                instructor_summary[key]['class_dates'].add(date)
            
            # 会議回数を集計
            for instructor_id, name, date in meeting_results:
                key = str(instructor_id)
                if key not in instructor_summary:
                    instructor_summary[key] = {'class_dates': set(), 'meeting_dates': set()}
                instructor_summary[key]['meeting_dates'].add(date)
            
            # すべての講師を取得してID順にソート
            all_instructors = self.db_manager.load_instructors_full()
            all_instructors_sorted = sorted(all_instructors, key=lambda x: int(x['instructor_id']))
            
            # monthlyフォルダの準備
            monthly_dir = "monthly"
            month_subdir = os.path.join(monthly_dir, month_str)
            old_dir = os.path.join(month_subdir, "old")
            
            if not os.path.exists(month_subdir):
                os.makedirs(month_subdir)
            if not os.path.exists(old_dir):
                os.makedirs(old_dir)
            
            csv_filename = os.path.join(month_subdir, f"【まとめ】出退勤記録_{month_str}.csv")
            
            # 既存ファイルがある場合はoldフォルダに移動
            self._move_to_old(csv_filename, old_dir, f"【まとめ】出退勤記録_{month_str}")
            
            # CSVファイルに書き込み
            with open(csv_filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(['【まとめ】講師ID', '【まとめ】講師名', '授業回数', '時間数（3.25h）', '会議回数', '出勤回数'])
                
                for instructor in all_instructors_sorted:
                    instructor_id = instructor['instructor_id']
                    name = instructor['name']
                    
                    summary = instructor_summary.get(instructor_id, {'class_dates': set(), 'meeting_dates': set()})
                    class_dates = summary['class_dates']
                    meeting_dates = summary['meeting_dates']
                    
                    class_count = len(class_dates)
                    meeting_count = len(meeting_dates)
                    total_attendance = len(class_dates | meeting_dates)
                    hours = class_count * 3.25
                    
                    writer.writerow([instructor_id, name, class_count, hours, meeting_count, total_attendance])
            
            # 統計情報
            total_instructors_registered = len(all_instructors_sorted)
            total_class_days = sum(len(summary.get('class_dates', set())) for summary in instructor_summary.values())
            total_meeting_days = sum(len(summary.get('meeting_dates', set())) for summary in instructor_summary.values())
            
            result = f"=== 統合月次集計エクスポート完了 ===\n\n"
            result += f"ファイル名: {csv_filename}\n"
            result += f"対象月: {month_str}\n\n"
            result += f"=== 集計結果 ===\n"
            result += f"登録講師数: {total_instructors_registered}人\n"
            result += f"総授業日数: {total_class_days}日\n"
            result += f"総会議日数: {total_meeting_days}日\n"
            result += f"総出勤日数: {total_class_days + total_meeting_days}日\n"
            
            return result
            
        except Exception as e:
            return f"統合月次集計エクスポートエラー: {e}"
    
    def _move_to_old(self, csv_filename, old_dir, base_name):
        """既存ファイルをoldフォルダに移動"""
        if os.path.exists(csv_filename):
            file_mtime = os.path.getmtime(csv_filename)
            file_datetime = datetime.fromtimestamp(file_mtime)
            timestamp_str = file_datetime.strftime("%H%M%S")
            
            old_filename = os.path.join(old_dir, f"{base_name}_{timestamp_str}.csv")
            
            if os.path.exists(old_filename):
                counter = 2
                while True:
                    old_filename = os.path.join(old_dir, f"{base_name}_{timestamp_str}_{counter}.csv")
                    if not os.path.exists(old_filename):
                        break
                    counter += 1
            
            shutil.move(csv_filename, old_filename)