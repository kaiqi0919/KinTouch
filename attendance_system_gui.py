# 出退勤管理システム（GUI版）
# pip install pyscard

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import time
import sqlite3
import csv
import os
import winsound
import shutil
import threading
from datetime import datetime, timezone, timedelta
from smartcard.System import readers
from smartcard.util import toHexString
from smartcard.Exceptions import CardConnectionException, NoCardException

# 日本時間のタイムゾーン設定
JST = timezone(timedelta(hours=9))

class AttendanceSystemGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("出退勤管理システム")
        self.root.geometry("600x500")
        
        # システムの初期化
        self.db_path = "attendance.db"
        self.instructors_csv = "instructors.csv"
        self.sound_enabled = True
        self.reader = None
        self.connection = None
        self.last_uid = None
        self.monitoring = False
        
        self.init_database()
        self.init_instructors_csv()
        self.initialize_reader()
        
        # メニュー画面を表示
        self.show_menu()
    
    def show_menu(self):
        """メニュー画面を表示"""
        # 既存のウィジェットをクリア
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # タイトル
        title_label = tk.Label(self.root, text="出退勤管理システム", font=("Arial", 20, "bold"))
        title_label.pack(pady=20)
        
        # ボタンフレーム
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)
        
        # ボタンのスタイル設定
        button_width = 20
        button_height = 2
        
        # 8つのメニューボタン
        buttons = [
            ("1. 打刻受付", self.show_attendance_monitor),
            ("2. 講師一覧", self.show_instructor_list),
            ("3. 打刻表示", self.show_attendance_records),
            ("4. 打刻サマリー", self.show_attendance_summary),
            ("5. 打刻修正", self.show_attendance_correction),
            ("6. CSVエクスポート", self.show_csv_export),
            ("7. 音量設定", self.toggle_sound_setting),
            ("8. 終了", self.exit_app)
        ]
        
        for i, (text, command) in enumerate(buttons):
            row = i // 2
            col = i % 2
            btn = tk.Button(button_frame, text=text, width=button_width, height=button_height,
                          font=("Arial", 12), command=command)
            btn.grid(row=row, column=col, padx=10, pady=10)
        
        # 音声状態表示
        sound_status = "有効" if self.sound_enabled else "無効"
        status_label = tk.Label(self.root, text=f"音声: {sound_status}", font=("Arial", 10))
        status_label.pack(side=tk.BOTTOM, pady=10)
    
    def show_attendance_monitor(self):
        """打刻受付画面"""
        # 既存のウィジェットをクリア
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # タイトル
        title_label = tk.Label(self.root, text="打刻受付", font=("Arial", 18, "bold"))
        title_label.pack(pady=20)
        
        # 状態表示ラベル
        status_label = tk.Label(self.root, text="カードをかざしてください...", 
                               font=("Arial", 14), fg="blue")
        status_label.pack(pady=20)
        
        # 打刻情報表示フレーム
        info_frame = tk.Frame(self.root, relief=tk.RIDGE, borderwidth=2)
        info_frame.pack(pady=20, padx=20, fill=tk.BOTH, expand=True)
        
        # 情報ラベル
        self.info_labels = {
            'instructor_id': tk.Label(info_frame, text="", font=("Arial", 16)),
            'name': tk.Label(info_frame, text="", font=("Arial", 20, "bold")),
            'uid': tk.Label(info_frame, text="", font=("Arial", 12)),
            'timestamp': tk.Label(info_frame, text="", font=("Arial", 14)),
            'action': tk.Label(info_frame, text="", font=("Arial", 18, "bold"))
        }
        
        for label in self.info_labels.values():
            label.pack(pady=5)
        
        # 終了ボタン（右上）
        exit_btn = tk.Button(self.root, text="終了", command=self.stop_monitoring,
                           font=("Arial", 12), bg="red", fg="white")
        exit_btn.place(x=520, y=10)
        
        # 監視開始
        self.monitoring = True
        self.monitor_thread = threading.Thread(target=self.monitor_cards, 
                                              args=(status_label,), daemon=True)
        self.monitor_thread.start()
    
    def monitor_cards(self, status_label):
        """カード監視スレッド"""
        while self.monitoring:
            try:
                if self.is_card_present():
                    if self.connect_to_card():
                        uid = self.get_card_uid()
                        if uid and uid != self.last_uid:
                            self.play_beep("card_detected")
                            self.process_attendance(uid, status_label)
                            self.last_uid = uid
                        self.disconnect()
                else:
                    if self.last_uid:
                        self.last_uid = None
                
                time.sleep(0.5)
            except Exception as e:
                print(f"監視エラー: {e}")
                time.sleep(1)
    
    def process_attendance(self, uid, status_label):
        """打刻処理"""
        instructor_info = self.get_instructor_info_by_uid(uid)
        
        if not instructor_info:
            self.root.after(0, lambda: status_label.config(
                text="未登録のカードです", fg="red"))
            self.play_beep("error")
            self.root.after(2000, lambda: status_label.config(
                text="カードをかざしてください...", fg="blue"))
            return
        
        # 最後の記録を確認
        last_record = self.get_last_record(uid)
        if last_record is None or last_record["type"] == "OUT":
            record_type = "IN"
            action = "出勤"
            action_color = "green"
        else:
            record_type = "OUT"
            action = "退勤"
            action_color = "orange"
        
        # データベースに記録
        jst_now = datetime.now(JST)
        timestamp_str = jst_now.strftime("%Y-%m-%d %H:%M:%S")
        
        if self.record_attendance_to_db(uid, instructor_info['name'], 
                                       instructor_info['instructor_id'], 
                                       record_type, timestamp_str):
            # 画面に表示
            self.root.after(0, lambda: self.display_attendance_info(
                instructor_info['instructor_id'],
                instructor_info['name'],
                uid,
                timestamp_str,
                action,
                action_color,
                status_label
            ))
            self.play_beep("success")
        else:
            self.play_beep("error")
    
    def display_attendance_info(self, instructor_id, name, uid, timestamp, action, color, status_label):
        """打刻情報を1秒間表示"""
        status_label.config(text=f"{action}記録完了！", fg=color)
        
        self.info_labels['instructor_id'].config(text=f"講師番号: {instructor_id}")
        self.info_labels['name'].config(text=f"{name}")
        self.info_labels['uid'].config(text=f"UID: {uid}")
        self.info_labels['timestamp'].config(text=f"{timestamp}")
        self.info_labels['action'].config(text=f"【{action}】", fg=color)
        
        # 1秒後にクリア
        self.root.after(1000, lambda: self.clear_attendance_info(status_label))
    
    def clear_attendance_info(self, status_label):
        """打刻情報をクリア"""
        status_label.config(text="カードをかざしてください...", fg="blue")
        for label in self.info_labels.values():
            label.config(text="")
    
    def stop_monitoring(self):
        """監視停止"""
        self.monitoring = False
        self.disconnect()
        self.show_menu()
    
    def show_instructor_list(self):
        """講師一覧画面"""
        # 既存のウィジェットをクリア
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # タイトル
        title_label = tk.Label(self.root, text="講師一覧", font=("Arial", 18, "bold"))
        title_label.pack(pady=10)
        
        # 講師登録ボタン（右上）
        register_btn = tk.Button(self.root, text="講師登録", 
                               command=self.show_instructor_registration,
                               font=("Arial", 12), bg="green", fg="white")
        register_btn.place(x=480, y=10)
        
        # テーブルフレーム
        table_frame = tk.Frame(self.root)
        table_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        # スクロールバー
        scrollbar = tk.Scrollbar(table_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Treeview
        columns = ('講師番号', 'カードUID', '講師名', '登録日時')
        tree = ttk.Treeview(table_frame, columns=columns, show='headings',
                           yscrollcommand=scrollbar.set)
        
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=140)
        
        # データ読み込み
        instructors = self.load_instructors_full()
        for instructor in instructors:
            tree.insert('', tk.END, values=(
                instructor['instructor_id'],
                instructor['card_uid'],
                instructor['name'],
                instructor['created_at']
            ))
        
        tree.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=tree.yview)
        
        # 件数表示
        count_label = tk.Label(self.root, text=f"登録講師数: {len(instructors)}人",
                             font=("Arial", 12))
        count_label.pack(pady=5)
        
        # 戻るボタン
        back_btn = tk.Button(self.root, text="戻る", command=self.show_menu,
                           font=("Arial", 12))
        back_btn.pack(pady=10)
    
    def show_instructor_registration(self):
        """講師登録画面"""
        # 新しいウィンドウを作成
        reg_window = tk.Toplevel(self.root)
        reg_window.title("講師登録")
        reg_window.geometry("400x300")
        
        tk.Label(reg_window, text="講師登録", font=("Arial", 16, "bold")).pack(pady=10)
        
        status_label = tk.Label(reg_window, text="カードをかざしてください...",
                              font=("Arial", 12), fg="blue")
        status_label.pack(pady=10)
        
        # 情報入力フレーム
        input_frame = tk.Frame(reg_window)
        input_frame.pack(pady=10)
        
        tk.Label(input_frame, text="カードUID:").grid(row=0, column=0, padx=5, pady=5)
        uid_entry = tk.Entry(input_frame, width=30, state='readonly')
        uid_entry.grid(row=0, column=1, padx=5, pady=5)
        
        next_id = self.get_next_instructor_id()
        tk.Label(input_frame, text="講師番号:").grid(row=1, column=0, padx=5, pady=5)
        id_entry = tk.Entry(input_frame, width=30)
        id_entry.insert(0, str(next_id))
        id_entry.grid(row=1, column=1, padx=5, pady=5)
        
        tk.Label(input_frame, text="講師名:").grid(row=2, column=0, padx=5, pady=5)
        name_entry = tk.Entry(input_frame, width=30)
        name_entry.grid(row=2, column=1, padx=5, pady=5)
        
        # カード検出用の変数
        detected_uid = {'uid': None}
        
        def check_card():
            """カード検出チェック"""
            if reg_window.winfo_exists():
                if self.is_card_present():
                    if self.connect_to_card():
                        uid = self.get_card_uid()
                        if uid and uid != detected_uid['uid']:
                            detected_uid['uid'] = uid
                            uid_entry.config(state='normal')
                            uid_entry.delete(0, tk.END)
                            uid_entry.insert(0, uid)
                            uid_entry.config(state='readonly')
                            status_label.config(text="カード検出！講師情報を入力してください", fg="green")
                            self.play_beep("card_detected")
                        self.disconnect()
                reg_window.after(500, check_card)
        
        check_card()
        
        def register():
            """登録実行"""
            uid = uid_entry.get()
            instructor_id = id_entry.get()
            name = name_entry.get()
            
            if not uid:
                messagebox.showerror("エラー", "カードをかざしてください")
                return
            
            if not name:
                messagebox.showerror("エラー", "講師名を入力してください")
                return
            
            try:
                instructor_id = int(instructor_id)
            except ValueError:
                messagebox.showerror("エラー", "講師番号は数値で入力してください")
                return
            
            if self.add_instructor_with_id(instructor_id, uid, name):
                messagebox.showinfo("成功", "講師を登録しました")
                reg_window.destroy()
                self.show_instructor_list()
            else:
                messagebox.showerror("エラー", "登録に失敗しました")
        
        # ボタン
        btn_frame = tk.Frame(reg_window)
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="登録", command=register, 
                 font=("Arial", 12), bg="green", fg="white", width=10).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="キャンセル", command=reg_window.destroy,
                 font=("Arial", 12), width=10).pack(side=tk.LEFT, padx=5)
    
    def show_attendance_records(self):
        """打刻表示画面"""
        # 既存のウィジェットをクリア
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # タイトル
        title_label = tk.Label(self.root, text="打刻表示", font=("Arial", 18, "bold"))
        title_label.pack(pady=10)
        
        # 日付入力フレーム
        input_frame = tk.Frame(self.root)
        input_frame.pack(pady=10)
        
        tk.Label(input_frame, text="日付 (YYYY-MM-DD):", font=("Arial", 12)).pack(side=tk.LEFT, padx=5)
        date_entry = tk.Entry(input_frame, width=15, font=("Arial", 12))
        date_entry.pack(side=tk.LEFT, padx=5)
        
        # 今日の日付をプレースホルダーとして表示
        today = datetime.now(JST).strftime("%Y-%m-%d")
        date_entry.insert(0, today)
        
        # テーブルフレーム
        table_frame = tk.Frame(self.root)
        table_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        # スクロールバー
        scrollbar = tk.Scrollbar(table_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Treeview
        columns = ('時刻', '講師名', '種別')
        tree = ttk.Treeview(table_frame, columns=columns, show='headings',
                           yscrollcommand=scrollbar.set)
        
        tree.heading('時刻', text='時刻')
        tree.heading('講師名', text='講師名')
        tree.heading('種別', text='種別')
        
        tree.column('時刻', width=150)
        tree.column('講師名', width=200)
        tree.column('種別', width=100)
        
        tree.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=tree.yview)
        
        # 件数表示
        count_label = tk.Label(self.root, text="", font=("Arial", 12))
        count_label.pack(pady=5)
        
        def display_records():
            """記録を表示"""
            date_str = date_entry.get().strip()
            if not date_str:
                date_str = today
            
            # 日付検証
            try:
                datetime.strptime(date_str, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("エラー", "日付形式が正しくありません (YYYY-MM-DD)")
                return
            
            # データ取得
            records = self.get_date_records(date_str)
            
            # テーブルクリア
            for item in tree.get_children():
                tree.delete(item)
            
            # データ挿入
            for name, record_type, timestamp in records:
                action = "出勤" if record_type == "IN" else "退勤"
                time_part = timestamp.split()[1] if ' ' in timestamp else timestamp
                tree.insert('', tk.END, values=(time_part, name, action))
            
            count_label.config(text=f"記録数: {len(records)}件")
        
        # 表示ボタン
        tk.Button(input_frame, text="表示", command=display_records,
                 font=("Arial", 12), bg="blue", fg="white").pack(side=tk.LEFT, padx=5)
        
        # 初期表示
        display_records()
        
        # 戻るボタン
        back_btn = tk.Button(self.root, text="戻る", command=self.show_menu,
                           font=("Arial", 12))
        back_btn.pack(pady=10)
    
    def show_attendance_summary(self):
        """打刻サマリー画面"""
        # 既存のウィジェットをクリア
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # タイトル
        title_label = tk.Label(self.root, text="打刻サマリー", font=("Arial", 18, "bold"))
        title_label.pack(pady=10)
        
        # 日付入力フレーム
        input_frame = tk.Frame(self.root)
        input_frame.pack(pady=10)
        
        tk.Label(input_frame, text="日付 (YYYY-MM-DD):", font=("Arial", 12)).pack(side=tk.LEFT, padx=5)
        date_entry = tk.Entry(input_frame, width=15, font=("Arial", 12))
        date_entry.pack(side=tk.LEFT, padx=5)
        
        # 今日の日付をプレースホルダーとして表示
        today = datetime.now(JST).strftime("%Y-%m-%d")
        date_entry.insert(0, today)
        
        # テーブルフレーム
        table_frame = tk.Frame(self.root)
        table_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        # スクロールバー
        scrollbar = tk.Scrollbar(table_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Treeview
        columns = ('講師名', '状態', '最終打刻時刻', '記録')
        tree = ttk.Treeview(table_frame, columns=columns, show='headings',
                           yscrollcommand=scrollbar.set)
        
        for col in columns:
            tree.heading(col, text=col)
        
        tree.column('講師名', width=120)
        tree.column('状態', width=80)
        tree.column('最終打刻時刻', width=150)
        tree.column('記録', width=200)
        
        tree.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=tree.yview)
        
        # 件数表示
        count_label = tk.Label(self.root, text="", font=("Arial", 12))
        count_label.pack(pady=5)
        
        def display_summary():
            """サマリーを表示"""
            date_str = date_entry.get().strip()
            if not date_str:
                date_str = today
            
            # 日付検証
            try:
                datetime.strptime(date_str, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("エラー", "日付形式が正しくありません (YYYY-MM-DD)")
                return
            
            # データ取得
            summary = self.get_date_summary(date_str)
            
            # テーブルクリア
            for item in tree.get_children():
                tree.delete(item)
            
            # データ挿入
            for name, status, last_time, record_str in summary:
                tree.insert('', tk.END, values=(name, status, last_time, record_str))
            
            count_label.config(text=f"打刻した講師数: {len(summary)}人")
        
        # 表示ボタン
        tk.Button(input_frame, text="表示", command=display_summary,
                 font=("Arial", 12), bg="blue", fg="white").pack(side=tk.LEFT, padx=5)
        
        # 初期表示
        display_summary()
        
        # 戻るボタン
        back_btn = tk.Button(self.root, text="戻る", command=self.show_menu,
                           font=("Arial", 12))
        back_btn.pack(pady=10)
    
    def show_attendance_correction(self):
        """打刻修正画面（パスワード認証付き）"""
        # パスワード入力ダイアログ
        password = simpledialog.askstring("パスワード入力", "パスワードを入力してください:", show='*')
        
        if password != "vgu2H8":
            if password is not None:  # キャンセルでない場合
                messagebox.showerror("エラー", "パスワードが正しくありません")
            return
        
        # 既存のウィジェットをクリア
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # タイトル
        title_label = tk.Label(self.root, text="打刻修正", font=("Arial", 18, "bold"))
        title_label.pack(pady=20)
        
        # 入力フレーム
        input_frame = tk.Frame(self.root)
        input_frame.pack(pady=20)
        
        # 講師番号入力
        tk.Label(input_frame, text="講師番号 (必須):", font=("Arial", 12)).grid(row=0, column=0, padx=10, pady=10, sticky='e')
        instructor_id_entry = tk.Entry(input_frame, width=20, font=("Arial", 12))
        instructor_id_entry.grid(row=0, column=1, padx=10, pady=10)
        
        # 時刻入力
        tk.Label(input_frame, text="時刻 (YYYY-MM-DD HH:MM:SS):", font=("Arial", 12)).grid(row=1, column=0, padx=10, pady=10, sticky='e')
        tk.Label(input_frame, text="※空欄の場合は現在時刻", font=("Arial", 9), fg="gray").grid(row=2, column=1, sticky='w')
        time_entry = tk.Entry(input_frame, width=30, font=("Arial", 12))
        time_entry.grid(row=1, column=1, padx=10, pady=10)
        
        # 打刻種別選択
        tk.Label(input_frame, text="打刻種別:", font=("Arial", 12)).grid(row=3, column=0, padx=10, pady=10, sticky='e')
        record_type_var = tk.StringVar(value="IN")
        tk.Radiobutton(input_frame, text="出勤", variable=record_type_var, value="IN", font=("Arial", 11)).grid(row=3, column=1, sticky='w')
        tk.Radiobutton(input_frame, text="退勤", variable=record_type_var, value="OUT", font=("Arial", 11)).grid(row=4, column=1, sticky='w')
        
        def register_correction():
            """修正を登録"""
            instructor_id = instructor_id_entry.get().strip()
            time_str = time_entry.get().strip()
            record_type = record_type_var.get()
            
            if not instructor_id:
                messagebox.showerror("エラー", "講師番号を入力してください")
                return
            
            try:
                instructor_id = int(instructor_id)
            except ValueError:
                messagebox.showerror("エラー", "講師番号は数値で入力してください")
                return
            
            # 講師情報取得
            instructor_info = self.get_instructor_info_by_id(instructor_id)
            if not instructor_info:
                messagebox.showerror("エラー", f"講師番号 {instructor_id} は登録されていません")
                return
            
            # 時刻設定
            if not time_str:
                timestamp = datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")
            else:
                # 時刻検証
                try:
                    datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")
                    timestamp = time_str
                except ValueError:
                    messagebox.showerror("エラー", "時刻形式が正しくありません (YYYY-MM-DD HH:MM:SS)")
                    return
            
            # データベースに記録
            if self.record_attendance_to_db(instructor_info['card_uid'], 
                                           instructor_info['name'],
                                           instructor_id,
                                           record_type,
                                           timestamp):
                action = "出勤" if record_type == "IN" else "退勤"
                messagebox.showinfo("成功", f"{instructor_info['name']} さんの{action}を記録しました\n時刻: {timestamp}")
                instructor_id_entry.delete(0, tk.END)
                time_entry.delete(0, tk.END)
            else:
                messagebox.showerror("エラー", "記録に失敗しました")
        
        # 登録ボタン
        tk.Button(self.root, text="登録", command=register_correction,
                 font=("Arial", 14), bg="green", fg="white", width=15, height=2).pack(pady=20)
        
        # 戻るボタン
        back_btn = tk.Button(self.root, text="戻る", command=self.show_menu,
                           font=("Arial", 12))
        back_btn.pack(pady=10)
    
    def show_csv_export(self):
        """CSVエクスポート画面"""
        # 既存のウィジェットをクリア
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # タイトル
        title_label = tk.Label(self.root, text="CSVエクスポート", font=("Arial", 18, "bold"))
        title_label.pack(pady=20)
        
        # 日付入力フレーム
        input_frame = tk.Frame(self.root)
        input_frame.pack(pady=20)
        
        tk.Label(input_frame, text="日付 (YYYY-MM-DD):", font=("Arial", 12)).pack(side=tk.LEFT, padx=5)
        date_entry = tk.Entry(input_frame, width=15, font=("Arial", 12))
        date_entry.pack(side=tk.LEFT, padx=5)
        
        # 今日の日付をプレースホルダーとして表示
        today = datetime.now(JST).strftime("%Y-%m-%d")
        date_entry.insert(0, today)
        
        # 結果表示エリア
        result_text = tk.Text(self.root, height=15, width=60, font=("Arial", 10))
        result_text.pack(pady=10, padx=20)
        
        def export_csv():
            """CSVエクスポート実行"""
            date_str = date_entry.get().strip()
            if not date_str:
                date_str = today
            
            # 日付検証
            try:
                datetime.strptime(date_str, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("エラー", "日付形式が正しくありません (YYYY-MM-DD)")
                return
            
            # エクスポート実行
            result = self.export_records_to_csv(date_str)
            
            # 結果表示
            result_text.delete(1.0, tk.END)
            result_text.insert(1.0, result)
        
        # エクスポートボタン
        tk.Button(input_frame, text="エクスポート", command=export_csv,
                 font=("Arial", 12), bg="green", fg="white").pack(side=tk.LEFT, padx=5)
        
        # 戻るボタン
        back_btn = tk.Button(self.root, text="戻る", command=self.show_menu,
                           font=("Arial", 12))
        back_btn.pack(pady=10)
    
    def toggle_sound_setting(self):
        """音量設定"""
        self.sound_enabled = not self.sound_enabled
        status = "有効" if self.sound_enabled else "無効"
        messagebox.showinfo("音量設定", f"音声機能を{status}にしました")
        
        # テスト音を再生
        if self.sound_enabled:
            self.play_beep("success")
    
    def exit_app(self):
        """アプリケーション終了"""
        if messagebox.askyesno("確認", "アプリケーションを終了しますか？"):
            self.monitoring = False
            self.disconnect()
            self.root.quit()
    
    # ===== データベース関連メソッド =====
    
    def init_database(self):
        """データベース初期化"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute("PRAGMA table_info(time_records)")
            columns = cursor.fetchall()
            column_names = [col[1] for col in columns]
            
            if 'instructor_name' not in column_names and columns:
                cursor.execute("ALTER TABLE time_records ADD COLUMN instructor_name TEXT DEFAULT '未登録'")
                cursor.execute("UPDATE time_records SET instructor_name = '未登録' WHERE instructor_name IS NULL")
            
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
            
        except Exception as e:
            print(f"データベース初期化エラー: {e}")
    
    def init_instructors_csv(self):
        """講師マスタCSV初期化"""
        if not os.path.exists(self.instructors_csv):
            try:
                with open(self.instructors_csv, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(["instructor_id", "card_uid", "name", "created_at"])
            except Exception as e:
                print(f"講師マスタCSV初期化エラー: {e}")
    
    def load_instructors(self):
        """CSVから講師データ読み込み（UID→名前の辞書）"""
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
    
    def load_instructors_full(self):
        """CSVから講師データ読み込み（全情報）"""
        instructors = []
        try:
            if os.path.exists(self.instructors_csv):
                with open(self.instructors_csv, 'r', encoding='utf-8') as csvfile:
                    reader = csv.DictReader(csvfile)
                    for row in reader:
                        instructors.append(row)
            return instructors
        except Exception as e:
            print(f"講師データ読み込みエラー: {e}")
            return []
    
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
    
    def get_instructor_info_by_uid(self, card_uid):
        """UIDから講師情報取得"""
        try:
            if not os.path.exists(self.instructors_csv):
                return None
            
            with open(self.instructors_csv, 'r', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    if row['card_uid'] == card_uid:
                        return {
                            'instructor_id': row['instructor_id'],
                            'name': row['name'],
                            'card_uid': row['card_uid']
                        }
            return None
        except Exception as e:
            print(f"講師情報取得エラー: {e}")
            return None
    
    def get_instructor_info_by_id(self, instructor_id):
        """講師番号から講師情報取得"""
        try:
            if not os.path.exists(self.instructors_csv):
                return None
            
            with open(self.instructors_csv, 'r', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    if int(row['instructor_id']) == instructor_id:
                        return {
                            'instructor_id': row['instructor_id'],
                            'name': row['name'],
                            'card_uid': row['card_uid']
                        }
            return None
        except Exception as e:
            print(f"講師情報取得エラー: {e}")
            return None
    
    def add_instructor_with_id(self, instructor_id, card_uid, name):
        """講師を指定IDで追加"""
        try:
            instructors = self.load_instructors()
            if card_uid in instructors:
                return False
            
            with open(self.instructors_csv, 'a', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                created_at = datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")
                writer.writerow([instructor_id, card_uid, name, created_at])
            
            return True
            
        except Exception as e:
            print(f"講師登録エラー: {e}")
            return False
    
    def get_last_record(self, card_uid):
        """最後の打刻記録を取得"""
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
    
    def record_attendance_to_db(self, card_uid, name, instructor_id, record_type, timestamp):
        """データベースに打刻記録"""
        max_retries = 3
        for attempt in range(max_retries):
            try:
                conn = sqlite3.connect(self.db_path, timeout=10.0)
                cursor = conn.cursor()
                
                cursor.execute("PRAGMA table_info(time_records)")
                columns = cursor.fetchall()
                column_names = [col[1] for col in columns]
                
                if 'instructor_id' in column_names:
                    cursor.execute(
                        "INSERT INTO time_records (instructor_id, card_uid, instructor_name, record_type, timestamp) VALUES (?, ?, ?, ?, ?)",
                        (instructor_id, card_uid, name, record_type, timestamp)
                    )
                else:
                    cursor.execute(
                        "INSERT INTO time_records (card_uid, instructor_name, record_type, timestamp) VALUES (?, ?, ?, ?)",
                        (card_uid, name, record_type, timestamp)
                    )
                
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
    
    def get_date_records(self, date_str):
        """特定日付の打刻記録取得"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT instructor_name, record_type, timestamp 
                FROM time_records
                WHERE DATE(timestamp) = ?
                ORDER BY timestamp DESC
            ''', (date_str,))
            
            results = cursor.fetchall()
            conn.close()
            return results
            
        except Exception as e:
            print(f"記録取得エラー: {e}")
            return []
    
    def get_date_summary(self, date_str):
        """特定日付のサマリー取得"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT instructor_name, record_type, timestamp
                FROM time_records
                WHERE DATE(timestamp) = ?
                ORDER BY instructor_name, timestamp
            ''', (date_str,))
            
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
    
    def export_records_to_csv(self, date_str):
        """CSVエクスポート"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT instructor_name, record_type, timestamp, card_uid
                FROM time_records
                WHERE DATE(timestamp) = ?
                ORDER BY timestamp
            ''', (date_str,))
            
            results = cursor.fetchall()
            conn.close()
            
            if not results:
                return f"{date_str} の打刻記録はありません。"
            
            # CSVファイル名を生成
            csv_filename = self.generate_unique_csv_filename(date_str)
            
            # CSVファイルに書き込み
            with open(csv_filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(['講師名', '打刻種別', '打刻日時', 'カードUID'])
                
                for name, record_type, timestamp, card_uid in results:
                    record_type_jp = "出勤" if record_type == "IN" else "退勤"
                    writer.writerow([name, record_type_jp, timestamp, card_uid])
            
            # NASへコピー
            nas_result = self.copy_to_nas(csv_filename)
            
            # 統計情報
            instructor_count = len(set(name for name, _, _, _ in results))
            in_count = sum(1 for _, record_type, _, _ in results if record_type == "IN")
            out_count = sum(1 for _, record_type, _, _ in results if record_type == "OUT")
            
            result = f"=== CSVエクスポート完了 ===\n\n"
            result += f"ファイル名: {csv_filename}\n"
            result += f"対象日: {date_str}\n"
            result += f"エクスポート件数: {len(results)}件\n\n"
            result += f"=== エクスポート内容概要 ===\n"
            result += f"打刻した講師数: {instructor_count}人\n"
            result += f"出勤記録: {in_count}件\n"
            result += f"退勤記録: {out_count}件\n"
            result += f"合計記録数: {len(results)}件\n\n"
            result += nas_result
            
            return result
            
        except Exception as e:
            return f"CSVエクスポートエラー: {e}"
    
    def generate_unique_csv_filename(self, date_str):
        """重複しないCSVファイル名を生成"""
        log_dir = "log"
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        base_filename = f"attendance_records_{date_str}"
        csv_filename = os.path.join(log_dir, f"{base_filename}.csv")
        
        if not os.path.exists(csv_filename):
            return csv_filename
        
        counter = 2
        while True:
            csv_filename = os.path.join(log_dir, f"{base_filename}_{counter}.csv")
            if not os.path.exists(csv_filename):
                return csv_filename
            counter += 1
    
    def copy_to_nas(self, csv_filename):
        """NASへコピー"""
        nas_path = r"\\NASTokyo\◆東京講師\NASsetup\kintouch_log"
        
        try:
            if not os.path.exists(nas_path):
                return f"警告: NASパス '{nas_path}' にアクセスできません。\nローカルのみに保存されました。"
            
            filename = os.path.basename(csv_filename)
            nas_file_path = os.path.join(nas_path, filename)
            
            shutil.copy2(csv_filename, nas_file_path)
            return f"NASへのコピー完了: {nas_file_path}"
            
        except PermissionError:
            return f"警告: NASパス '{nas_path}' への書き込み権限がありません。\nローカルのみに保存されました。"
        except Exception as e:
            return f"警告: NASへのコピー中にエラーが発生しました: {e}\nローカルのみに保存されました。"
    
    # ===== カードリーダー関連メソッド =====
    
    def initialize_reader(self):
        """リーダー初期化"""
        try:
            r = readers()
            if len(r) == 0:
                messagebox.showwarning("警告", "カードリーダーが見つかりません")
                return False
            
            self.reader = r[0]
            return True
            
        except Exception as e:
            messagebox.showerror("エラー", f"リーダー初期化エラー: {e}")
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
            
        except (NoCardException, CardConnectionException):
            return False
        except Exception:
            return False
    
    def get_card_uid(self):
        """カードUID取得"""
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
        """接続切断"""
        if self.connection:
            try:
                self.connection.disconnect()
            except:
                pass
            self.connection = None
    
    def is_card_present(self):
        """カード存在チェック"""
        try:
            test_connection = self.reader.createConnection()
            test_connection.connect()
            test_connection.disconnect()
            return True
        except (NoCardException, CardConnectionException):
            return False
        except Exception:
            return False
    
    # ===== 音声関連メソッド =====
    
    def play_beep(self, beep_type="success"):
        """ビープ音再生"""
        if not self.sound_enabled:
            return
        
        try:
            if beep_type == "success":
                winsound.Beep(1000, 200)
                time.sleep(0.1)
                winsound.Beep(1000, 200)
            elif beep_type == "error":
                winsound.Beep(400, 500)
            elif beep_type == "card_detected":
                winsound.Beep(800, 150)
        except Exception as e:
            print(f"音声再生エラー: {e}")

def main():
    """メイン関数"""
    root = tk.Tk()
    app = AttendanceSystemGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()