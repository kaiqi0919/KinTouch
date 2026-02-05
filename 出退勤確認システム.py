# 出退勤管理システム（GUI版・複数リーダー対応）
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
import json
import hashlib
from datetime import datetime, timezone, timedelta
from smartcard.System import readers
from smartcard.util import toHexString
from smartcard.Exceptions import CardConnectionException, NoCardException

# 日本時間のタイムゾーン設定
JST = timezone(timedelta(hours=9))

# パスワードのハッシュ値（SHA-256）
PASSWORD_HASH = "944381cba581a7ee3f59b5e2a97686c9b48fc0b7da14eb3fceff5f84f62d0ff7"

class AttendanceSystemGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("出退勤確認システム（複数リーダー対応）")
        
        # ウィンドウを画面中央に配置
        window_width = 900
        window_height = 600
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        center_x = int((screen_width - window_width) / 2)
        center_y = int((screen_height - window_height) / 2)
        self.root.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        
        # システムの初期化
        self.data_dir = "data"
        if not os.path.exists(self.data_dir):
            os.makedirs(self.data_dir)
        
        self.db_path = os.path.join(self.data_dir, "attendance.db")
        self.instructors_csv = os.path.join(self.data_dir, "instructors.csv")
        self.config_path = "reader_config.json"
        self.sound_enabled = True
        
        # リーダー関連
        self.class_reader = None
        self.meeting_reader = None
        self.class_connection = None
        self.meeting_connection = None
        self.last_class_uid = None
        self.last_meeting_uid = None
        self.monitoring = False
        self.clear_timer_class = None
        self.clear_timer_meeting = None
        
        self.init_database()
        self.init_instructors_csv()
        
        # 設定の読み込みと初期化
        if not self.load_config():
            self.show_reader_setup()
        else:
            if not self.initialize_readers():
                messagebox.showwarning("警告", "リーダーの初期化に失敗しました\n設定を確認してください")
                self.show_reader_setup()
            else:
                self.show_menu()
    
    def load_config(self):
        """設定ファイルの読み込み"""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    if config.get('configured', False):
                        self.class_reader_name = config.get('class_reader', '')
                        self.meeting_reader_name = config.get('meeting_reader', '')
                        return True
            return False
        except Exception as e:
            print(f"設定読み込みエラー: {e}")
            return False
    
    def save_config(self, class_reader_name, meeting_reader_name):
        """設定ファイルの保存"""
        try:
            config = {
                'class_reader': class_reader_name,
                'meeting_reader': meeting_reader_name,
                'configured': True
            }
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"設定保存エラー: {e}")
            return False
    
    def show_reader_setup(self):
        """リーダー設定画面"""
        # 既存のウィジェットをクリア
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # タイトル
        title_label = tk.Label(self.root, text="リーダー設定", font=("Arial", 18, "bold"))
        title_label.pack(pady=20)
        
        # 説明
        info_label = tk.Label(self.root, text="接続されているリーダーを授業用と会議用に割り当ててください",
                             font=("Arial", 12))
        info_label.pack(pady=10)
        
        # リーダー一覧取得
        try:
            r = readers()
            if len(r) < 2:
                error_label = tk.Label(self.root, 
                    text=f"エラー: {len(r)}台のリーダーしか検出されていません\n2台のリーダーを接続してください",
                    font=("Arial", 12), fg="red")
                error_label.pack(pady=20)
                
                retry_btn = tk.Button(self.root, text="再試行", 
                    command=self.show_reader_setup, font=("Arial", 12))
                retry_btn.pack(pady=10)
                return
            
            reader_names = [reader.name for reader in r]
            
            # 選択フレーム
            select_frame = tk.Frame(self.root)
            select_frame.pack(pady=20)
            
            # 授業用リーダー選択
            tk.Label(select_frame, text="授業用リーダー:", 
                    font=("Arial", 12)).grid(row=0, column=0, padx=10, pady=10, sticky='e')
            class_combo = ttk.Combobox(select_frame, values=reader_names, 
                                      width=40, font=("Arial", 10), state='readonly')
            class_combo.grid(row=0, column=1, padx=10, pady=10)
            if reader_names:
                class_combo.set(reader_names[0])
            
            # 会議用リーダー選択
            tk.Label(select_frame, text="会議用リーダー:", 
                    font=("Arial", 12)).grid(row=1, column=0, padx=10, pady=10, sticky='e')
            meeting_combo = ttk.Combobox(select_frame, values=reader_names, 
                                        width=40, font=("Arial", 10), state='readonly')
            meeting_combo.grid(row=1, column=1, padx=10, pady=10)
            if len(reader_names) > 1:
                meeting_combo.set(reader_names[1])
            
            def save_and_continue():
                class_reader = class_combo.get()
                meeting_reader = meeting_combo.get()
                
                if not class_reader or not meeting_reader:
                    messagebox.showerror("エラー", "両方のリーダーを選択してください")
                    return
                
                if class_reader == meeting_reader:
                    messagebox.showerror("エラー", "異なるリーダーを選択してください")
                    return
                
                if self.save_config(class_reader, meeting_reader):
                    self.class_reader_name = class_reader
                    self.meeting_reader_name = meeting_reader
                    if self.initialize_readers():
                        messagebox.showinfo("成功", "リーダー設定が完了しました")
                        self.show_menu()
                    else:
                        messagebox.showerror("エラー", "リーダーの初期化に失敗しました")
                else:
                    messagebox.showerror("エラー", "設定の保存に失敗しました")
            
            # 保存ボタン
            tk.Button(self.root, text="設定を保存", command=save_and_continue,
                     font=("Arial", 14), bg="green", fg="white", width=15).pack(pady=20)
            
        except Exception as e:
            error_label = tk.Label(self.root, 
                text=f"エラー: {e}\nリーダーの接続を確認してください",
                font=("Arial", 12), fg="red")
            error_label.pack(pady=20)
    
    def initialize_readers(self):
        """リーダー初期化"""
        try:
            r = readers()
            if len(r) < 2:
                return False
            
            # リーダーを名前で識別
            for reader in r:
                if reader.name == self.class_reader_name:
                    self.class_reader = reader
                elif reader.name == self.meeting_reader_name:
                    self.meeting_reader = reader
            
            if self.class_reader is None or self.meeting_reader is None:
                return False
            
            return True
            
        except Exception as e:
            print(f"リーダー初期化エラー: {e}")
            return False
    
    def show_menu(self):
        """メニュー画面を表示"""
        # 既存のウィジェットをクリア
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # タイトル
        title_label = tk.Label(self.root, text="出退勤確認システム", font=("Arial", 20, "bold"))
        title_label.pack(pady=20)
        
        # ボタンフレーム
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)
        
        # ボタンのスタイル設定
        button_width = 20
        button_height = 2
        
        # メニューボタン
        buttons = [
            ("1. 打刻受付", self.show_attendance_monitor),
            ("2. 講師一覧", self.show_instructor_list),
            ("3. 打刻表示", self.show_attendance_records),
            ("4. 打刻サマリー", self.show_attendance_summary),
            ("5. 打刻修正", self.show_attendance_correction),
            ("6. 音量設定", self.toggle_sound_setting),
            ("7. 日次集計", self.show_csv_export),
            ("8. 月次集計", self.show_monthly_summary),
            ("9. リーダー設定", self.show_reader_setup),
            ("0. 終了", self.exit_app)
        ]
        
        for i, (text, command) in enumerate(buttons):
            row = i // 3
            col = i % 3
            btn = tk.Button(button_frame, text=text, width=button_width, height=button_height,
                          font=("Arial", 12), command=command)
            btn.grid(row=row, column=col, padx=10, pady=10)
        
        # 音声状態表示
        sound_status = "有効" if self.sound_enabled else "無効"
        status_label = tk.Label(self.root, text=f"音声: {sound_status}", font=("Arial", 10))
        status_label.pack(side=tk.BOTTOM, pady=10)
    
    def show_attendance_monitor(self):
        """打刻受付画面（2分割）"""
        # 既存のウィジェットをクリア
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # タイトル
        title_label = tk.Label(self.root, text="打刻受付", font=("Arial", 18, "bold"))
        title_label.pack(pady=10)
        
        # メインフレーム（2分割）
        main_frame = tk.Frame(self.root)
        main_frame.pack(pady=10, fill=tk.BOTH, expand=True)
        
        # 左側：授業用
        class_frame = tk.Frame(main_frame, relief=tk.RIDGE, borderwidth=2)
        class_frame.pack(side=tk.LEFT, padx=10, fill=tk.BOTH, expand=True)
        
        tk.Label(class_frame, text="授業用", font=("Arial", 16, "bold"), 
                bg="lightblue").pack(fill=tk.X, pady=5)
        
        class_status = tk.Label(class_frame, text="カードをかざしてください...", 
                               font=("Arial", 12), fg="blue")
        class_status.pack(pady=10)
        
        class_info_frame = tk.Frame(class_frame)
        class_info_frame.pack(pady=10, fill=tk.BOTH, expand=True)
        
        self.class_info_labels = {
            'instructor_id': tk.Label(class_info_frame, text="", font=("Arial", 14)),
            'name': tk.Label(class_info_frame, text="", font=("Arial", 18, "bold")),
            'timestamp': tk.Label(class_info_frame, text="", font=("Arial", 12)),
            'action': tk.Label(class_info_frame, text="", font=("Arial", 16, "bold"))
        }
        
        for label in self.class_info_labels.values():
            label.pack(pady=3)
        
        # 右側：会議用
        meeting_frame = tk.Frame(main_frame, relief=tk.RIDGE, borderwidth=2)
        meeting_frame.pack(side=tk.RIGHT, padx=10, fill=tk.BOTH, expand=True)
        
        tk.Label(meeting_frame, text="会議用", font=("Arial", 16, "bold"), 
                bg="lightgreen").pack(fill=tk.X, pady=5)
        
        meeting_status = tk.Label(meeting_frame, text="カードをかざしてください...", 
                                 font=("Arial", 12), fg="blue")
        meeting_status.pack(pady=10)
        
        meeting_info_frame = tk.Frame(meeting_frame)
        meeting_info_frame.pack(pady=10, fill=tk.BOTH, expand=True)
        
        self.meeting_info_labels = {
            'instructor_id': tk.Label(meeting_info_frame, text="", font=("Arial", 14)),
            'name': tk.Label(meeting_info_frame, text="", font=("Arial", 18, "bold")),
            'timestamp': tk.Label(meeting_info_frame, text="", font=("Arial", 12)),
            'action': tk.Label(meeting_info_frame, text="", font=("Arial", 16, "bold"))
        }
        
        for label in self.meeting_info_labels.values():
            label.pack(pady=3)
        
        # 終了ボタン
        exit_btn = tk.Button(self.root, text="終了", command=self.stop_monitoring,
                           font=("Arial", 12), bg="red", fg="white")
        exit_btn.pack(pady=10)
        
        # 監視開始
        self.monitoring = True
        self.class_monitor_thread = threading.Thread(
            target=self.monitor_cards, 
            args=(self.class_reader, class_status, 'class'), 
            daemon=True
        )
        self.meeting_monitor_thread = threading.Thread(
            target=self.monitor_cards, 
            args=(self.meeting_reader, meeting_status, 'meeting'), 
            daemon=True
        )
        self.class_monitor_thread.start()
        self.meeting_monitor_thread.start()
    
    def monitor_cards(self, reader, status_label, reader_type):
        """カード監視スレッド"""
        connection = None
        last_uid = None
        
        while self.monitoring:
            try:
                if self.is_card_present(reader):
                    connection = self.connect_to_card(reader)
                    if connection:
                        uid = self.get_card_uid(connection)
                        if uid and uid != last_uid:
                            self.play_beep("card_detected")
                            self.process_attendance(uid, status_label, reader_type)
                            last_uid = uid
                        self.disconnect(connection)
                        connection = None
                else:
                    if last_uid:
                        last_uid = None
                
                time.sleep(0.5)
            except Exception as e:
                print(f"監視エラー ({reader_type}): {e}")
                time.sleep(1)
        
        if connection:
            self.disconnect(connection)
    
    def process_attendance(self, uid, status_label, reader_type):
        """打刻処理"""
        instructor_info = self.get_instructor_info_by_uid(uid)
        
        if not instructor_info:
            self.root.after(0, lambda: status_label.config(
                text="未登録のカードです", fg="red"))
            self.play_beep("error")
            self.root.after(2000, lambda: status_label.config(
                text="カードをかざしてください...", fg="blue"))
            return
        
        # テーブル名の決定
        table_name = "time_records" if reader_type == "class" else "meeting_records"
        
        # 最後の記録を確認
        last_record = self.get_last_record(uid, table_name)
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
                                       record_type, timestamp_str, table_name):
            # 画面に表示
            self.root.after(0, lambda: self.display_attendance_info(
                instructor_info['instructor_id'],
                instructor_info['name'],
                uid,
                timestamp_str,
                action,
                action_color,
                status_label,
                reader_type
            ))
            self.play_beep("success")
        else:
            self.play_beep("error")
    
    def display_attendance_info(self, instructor_id, name, uid, timestamp, action, color, status_label, reader_type):
        """打刻情報を3秒間表示"""
        # ラベルの選択
        if reader_type == "class":
            info_labels = self.class_info_labels
            timer_attr = 'clear_timer_class'
        else:
            info_labels = self.meeting_info_labels
            timer_attr = 'clear_timer_meeting'
        
        # 前回のタイマーがあればキャンセル
        timer_id = getattr(self, timer_attr, None)
        if timer_id is not None:
            self.root.after_cancel(timer_id)
        
        status_label.config(text=f"{action}記録完了！", fg=color)
        
        info_labels['instructor_id'].config(text=f"講師番号: {instructor_id}")
        info_labels['name'].config(text=f"{name}")
        info_labels['timestamp'].config(text=f"{timestamp}")
        info_labels['action'].config(text=f"【{action}】", fg=color)
        
        # 3秒後にクリア
        new_timer_id = self.root.after(3000, lambda: self.clear_attendance_info(status_label, info_labels))
        setattr(self, timer_attr, new_timer_id)
    
    def clear_attendance_info(self, status_label, info_labels):
        """打刻情報をクリア"""
        status_label.config(text="カードをかざしてください...", fg="blue")
        for label in info_labels.values():
            label.config(text="")
    
    def stop_monitoring(self):
        """監視停止"""
        self.monitoring = False
        time.sleep(0.5)  # スレッドの終了を待つ
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
        register_btn.place(x=680, y=10)
        
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
            tree.column(col, width=180)
        
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
        reg_window.geometry("450x350")
        
        tk.Label(reg_window, text="講師登録", font=("Arial", 16, "bold")).pack(pady=10)
        
        # リーダー選択
        reader_frame = tk.Frame(reg_window)
        reader_frame.pack(pady=5)
        
        tk.Label(reader_frame, text="使用するリーダー:", font=("Arial", 11)).pack(side=tk.LEFT, padx=5)
        reader_var = tk.StringVar(value="class")
        tk.Radiobutton(reader_frame, text="授業用", variable=reader_var, 
                      value="class", font=("Arial", 10)).pack(side=tk.LEFT)
        tk.Radiobutton(reader_frame, text="会議用", variable=reader_var, 
                      value="meeting", font=("Arial", 10)).pack(side=tk.LEFT)
        
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
                selected_reader = self.class_reader if reader_var.get() == "class" else self.meeting_reader
                
                if self.is_card_present(selected_reader):
                    connection = self.connect_to_card(selected_reader)
                    if connection:
                        uid = self.get_card_uid(connection)
                        if uid and uid != detected_uid['uid']:
                            detected_uid['uid'] = uid
                            uid_entry.config(state='normal')
                            uid_entry.delete(0, tk.END)
                            uid_entry.insert(0, uid)
                            uid_entry.config(state='readonly')
                            status_label.config(text="カード検出！講師情報を入力してください", fg="green")
                            self.play_beep("card_detected")
                        self.disconnect(connection)
                reg_window.after(500, check_card)
        
        check_card()
        
        def register():
            """登録実行"""
            uid = uid_entry.get().strip()
            instructor_id = id_entry.get().strip()
            name = name_entry.get().strip()
            
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
        """打刻表示画面（2分割）"""
        # 既存のウィジェットをクリア
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # タイトル
        title_label = tk.Label(self.root, text="打刻表示", font=("Arial", 18, "bold"))
        title_label.pack(pady=10)
        
        # 入力フレーム
        input_frame = tk.Frame(self.root)
        input_frame.pack(pady=10)
        
        tk.Label(input_frame, text="日付 (YYYY-MM-DD):", font=("Arial", 12)).pack(side=tk.LEFT, padx=5)
        date_entry = tk.Entry(input_frame, width=15, font=("Arial", 12))
        date_entry.pack(side=tk.LEFT, padx=5)
        
        # 今日の日付をプレースホルダーとして表示
        today = datetime.now(JST).strftime("%Y-%m-%d")
        date_entry.insert(0, today)
        
        # メインフレーム（2分割）
        main_frame = tk.Frame(self.root)
        main_frame.pack(pady=10, fill=tk.BOTH, expand=True, padx=10)
        
        # 左側：授業用
        class_frame = tk.Frame(main_frame, relief=tk.RIDGE, borderwidth=2)
        class_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        tk.Label(class_frame, text="授業用", font=("Arial", 14, "bold"), 
                bg="lightblue").pack(fill=tk.X, pady=5)
        
        # 授業用テーブル
        class_table_frame = tk.Frame(class_frame)
        class_table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        class_scrollbar = tk.Scrollbar(class_table_frame)
        class_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        columns = ('時刻', '講師名', '種別')
        class_tree = ttk.Treeview(class_table_frame, columns=columns, show='headings',
                                  yscrollcommand=class_scrollbar.set, height=15)
        
        class_tree.heading('時刻', text='時刻')
        class_tree.heading('講師名', text='講師名')
        class_tree.heading('種別', text='種別')
        
        class_tree.column('時刻', width=100)
        class_tree.column('講師名', width=120)
        class_tree.column('種別', width=60)
        
        class_tree.pack(fill=tk.BOTH, expand=True)
        class_scrollbar.config(command=class_tree.yview)
        
        class_count_label = tk.Label(class_frame, text="", font=("Arial", 10))
        class_count_label.pack(pady=5)
        
        # 右側：会議用
        meeting_frame = tk.Frame(main_frame, relief=tk.RIDGE, borderwidth=2)
        meeting_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)
        
        tk.Label(meeting_frame, text="会議用", font=("Arial", 14, "bold"), 
                bg="lightgreen").pack(fill=tk.X, pady=5)
        
        # 会議用テーブル
        meeting_table_frame = tk.Frame(meeting_frame)
        meeting_table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        meeting_scrollbar = tk.Scrollbar(meeting_table_frame)
        meeting_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        meeting_tree = ttk.Treeview(meeting_table_frame, columns=columns, show='headings',
                                    yscrollcommand=meeting_scrollbar.set, height=15)
        
        meeting_tree.heading('時刻', text='時刻')
        meeting_tree.heading('講師名', text='講師名')
        meeting_tree.heading('種別', text='種別')
        
        meeting_tree.column('時刻', width=100)
        meeting_tree.column('講師名', width=120)
        meeting_tree.column('種別', width=60)
        
        meeting_tree.pack(fill=tk.BOTH, expand=True)
        meeting_scrollbar.config(command=meeting_tree.yview)
        
        meeting_count_label = tk.Label(meeting_frame, text="", font=("Arial", 10))
        meeting_count_label.pack(pady=5)
        
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
            
            # 授業用データ取得
            class_records = self.get_date_records(date_str, "time_records")
            
            # 授業用テーブルクリア
            for item in class_tree.get_children():
                class_tree.delete(item)
            
            # 授業用データ挿入
            for name, record_type, timestamp in class_records:
                action = "出勤" if record_type == "IN" else "退勤"
                time_part = timestamp.split()[1] if ' ' in timestamp else timestamp
                class_tree.insert('', tk.END, values=(time_part, name, action))
            
            class_count_label.config(text=f"記録数: {len(class_records)}件")
            
            # 会議用データ取得
            meeting_records = self.get_date_records(date_str, "meeting_records")
            
            # 会議用テーブルクリア
            for item in meeting_tree.get_children():
                meeting_tree.delete(item)
            
            # 会議用データ挿入
            for name, record_type, timestamp in meeting_records:
                action = "出勤" if record_type == "IN" else "退勤"
                time_part = timestamp.split()[1] if ' ' in timestamp else timestamp
                meeting_tree.insert('', tk.END, values=(time_part, name, action))
            
            meeting_count_label.config(text=f"記録数: {len(meeting_records)}件")
        
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
        """打刻サマリー画面（2分割）"""
        # 既存のウィジェットをクリア
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # タイトル
        title_label = tk.Label(self.root, text="打刻サマリー", font=("Arial", 18, "bold"))
        title_label.pack(pady=10)
        
        # 入力フレーム
        input_frame = tk.Frame(self.root)
        input_frame.pack(pady=10)
        
        tk.Label(input_frame, text="日付 (YYYY-MM-DD):", font=("Arial", 12)).pack(side=tk.LEFT, padx=5)
        date_entry = tk.Entry(input_frame, width=15, font=("Arial", 12))
        date_entry.pack(side=tk.LEFT, padx=5)
        
        # 今日の日付をプレースホルダーとして表示
        today = datetime.now(JST).strftime("%Y-%m-%d")
        date_entry.insert(0, today)
        
        # メインフレーム（2分割）
        main_frame = tk.Frame(self.root)
        main_frame.pack(pady=10, fill=tk.BOTH, expand=True, padx=10)
        
        # 左側：授業用
        class_frame = tk.Frame(main_frame, relief=tk.RIDGE, borderwidth=2)
        class_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        tk.Label(class_frame, text="授業用", font=("Arial", 14, "bold"), 
                bg="lightblue").pack(fill=tk.X, pady=5)
        
        # 授業用テーブル
        class_table_frame = tk.Frame(class_frame)
        class_table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        class_scrollbar = tk.Scrollbar(class_table_frame)
        class_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        columns = ('講師名', '状態', '最終打刻時刻', '記録')
        class_tree = ttk.Treeview(class_table_frame, columns=columns, show='headings',
                                  yscrollcommand=class_scrollbar.set, height=15)
        
        for col in columns:
            class_tree.heading(col, text=col)
        
        class_tree.column('講師名', width=80)
        class_tree.column('状態', width=60)
        class_tree.column('最終打刻時刻', width=120)
        class_tree.column('記録', width=150)
        
        class_tree.pack(fill=tk.BOTH, expand=True)
        class_scrollbar.config(command=class_tree.yview)
        
        class_count_label = tk.Label(class_frame, text="", font=("Arial", 10))
        class_count_label.pack(pady=5)
        
        # 右側：会議用
        meeting_frame = tk.Frame(main_frame, relief=tk.RIDGE, borderwidth=2)
        meeting_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)
        
        tk.Label(meeting_frame, text="会議用", font=("Arial", 14, "bold"), 
                bg="lightgreen").pack(fill=tk.X, pady=5)
        
        # 会議用テーブル
        meeting_table_frame = tk.Frame(meeting_frame)
        meeting_table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        meeting_scrollbar = tk.Scrollbar(meeting_table_frame)
        meeting_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        meeting_tree = ttk.Treeview(meeting_table_frame, columns=columns, show='headings',
                                    yscrollcommand=meeting_scrollbar.set, height=15)
        
        for col in columns:
            meeting_tree.heading(col, text=col)
        
        meeting_tree.column('講師名', width=80)
        meeting_tree.column('状態', width=60)
        meeting_tree.column('最終打刻時刻', width=120)
        meeting_tree.column('記録', width=150)
        
        meeting_tree.pack(fill=tk.BOTH, expand=True)
        meeting_scrollbar.config(command=meeting_tree.yview)
        
        meeting_count_label = tk.Label(meeting_frame, text="", font=("Arial", 10))
        meeting_count_label.pack(pady=5)
        
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
            
            # 授業用データ取得
            class_summary = self.get_date_summary(date_str, "time_records")
            
            # 授業用テーブルクリア
            for item in class_tree.get_children():
                class_tree.delete(item)
            
            # 授業用データ挿入
            for name, status, last_time, record_str in class_summary:
                class_tree.insert('', tk.END, values=(name, status, last_time, record_str))
            
            class_count_label.config(text=f"打刻した講師数: {len(class_summary)}人")
            
            # 会議用データ取得
            meeting_summary = self.get_date_summary(date_str, "meeting_records")
            
            # 会議用テーブルクリア
            for item in meeting_tree.get_children():
                meeting_tree.delete(item)
            
            # 会議用データ挿入
            for name, status, last_time, record_str in meeting_summary:
                meeting_tree.insert('', tk.END, values=(name, status, last_time, record_str))
            
            meeting_count_label.config(text=f"打刻した講師数: {len(meeting_summary)}人")
        
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
        
        # パスワードをハッシュ化して検証
        if password is None:
            return
        
        password_hash = hashlib.sha256(password.encode('utf-8')).hexdigest()
        if password_hash != PASSWORD_HASH:
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
        
        # テーブル選択
        tk.Label(input_frame, text="種別:", font=("Arial", 12)).grid(row=0, column=0, padx=10, pady=5, sticky='e')
        table_var = tk.StringVar(value="time_records")
        type_frame = tk.Frame(input_frame)
        type_frame.grid(row=0, column=1, padx=10, pady=5, sticky='w')
        tk.Radiobutton(type_frame, text="授業用", variable=table_var, 
                      value="time_records", font=("Arial", 11)).pack(side=tk.LEFT)
        tk.Radiobutton(type_frame, text="会議用", variable=table_var, 
                      value="meeting_records", font=("Arial", 11)).pack(side=tk.LEFT)
        
        # 講師選択
        tk.Label(input_frame, text="講師選択 (必須):", font=("Arial", 12)).grid(row=1, column=0, padx=10, pady=10, sticky='e')
        
        instructors = self.load_instructors_full()
        instructor_options = [f"{inst['instructor_id']}: {inst['name']}" for inst in instructors]
        
        instructor_combo = ttk.Combobox(input_frame, values=instructor_options, width=28, font=("Arial", 11), state='readonly')
        instructor_combo.grid(row=1, column=1, padx=10, pady=10)
        if instructor_options:
            instructor_combo.set("講師を選択してください")
        
        # 時刻入力
        tk.Label(input_frame, text="時刻 (YYYY-MM-DD HH:MM:SS):", font=("Arial", 12)).grid(row=2, column=0, padx=10, pady=10, sticky='e')
        tk.Label(input_frame, text="※空欄の場合は現在時刻", font=("Arial", 9), fg="gray").grid(row=3, column=1, sticky='w')
        time_entry = tk.Entry(input_frame, width=30, font=("Arial", 12))
        time_entry.grid(row=2, column=1, padx=10, pady=10)
        
        # 打刻種別選択
        tk.Label(input_frame, text="打刻種別:", font=("Arial", 12)).grid(row=4, column=0, padx=10, pady=10, sticky='e')
        record_type_var = tk.StringVar(value="IN")
        tk.Radiobutton(input_frame, text="出勤", variable=record_type_var, value="IN", font=("Arial", 11)).grid(row=4, column=1, sticky='w')
        tk.Radiobutton(input_frame, text="退勤", variable=record_type_var, value="OUT", font=("Arial", 11)).grid(row=5, column=1, sticky='w')
        
        def register_correction():
            """修正を登録"""
            selected = instructor_combo.get()
            time_str = time_entry.get().strip()
            record_type = record_type_var.get()
            table_name = table_var.get()
            
            if not selected or selected == "講師を選択してください":
                messagebox.showerror("エラー", "講師を選択してください")
                return
            
            try:
                instructor_id = int(selected.split(':')[0])
            except (ValueError, IndexError):
                messagebox.showerror("エラー", "講師の選択が正しくありません")
                return
            
            instructor_info = self.get_instructor_info_by_id(instructor_id)
            if not instructor_info:
                messagebox.showerror("エラー", f"講師番号 {instructor_id} は登録されていません")
                return
            
            if not time_str:
                timestamp = datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")
            else:
                try:
                    datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")
                    timestamp = time_str
                except ValueError:
                    messagebox.showerror("エラー", "時刻形式が正しくありません (YYYY-MM-DD HH:%M:%S)")
                    return
            
            if self.record_attendance_to_db(instructor_info['card_uid'], 
                                           instructor_info['name'],
                                           instructor_id,
                                           record_type,
                                           timestamp,
                                           table_name):
                action = "出勤" if record_type == "IN" else "退勤"
                table_type = "授業用" if table_name == "time_records" else "会議用"
                messagebox.showinfo("成功", f"{table_type}\n{instructor_info['name']} さんの{action}を記録しました\n時刻: {timestamp}")
                instructor_combo.set("講師を選択してください")
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
        """日次集計画面"""
        # 既存のウィジェットをクリア
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # タイトル
        title_label = tk.Label(self.root, text="日次集計", font=("Arial", 18, "bold"))
        title_label.pack(pady=20)
        
        # 入力フレーム
        input_frame = tk.Frame(self.root)
        input_frame.pack(pady=20)
        
        # 対象日入力
        tk.Label(input_frame, text="対象日 (YYYY-MM-DD):", font=("Arial", 12)).pack(side=tk.LEFT, padx=5)
        date_entry = tk.Entry(input_frame, width=15, font=("Arial", 12))
        date_entry.pack(side=tk.LEFT, padx=5)
        
        # 今日の日付をプレースホルダーとして表示
        today = datetime.now(JST).strftime("%Y-%m-%d")
        date_entry.insert(0, today)
        
        # オプションフレーム
        option_frame = tk.Frame(self.root)
        option_frame.pack(pady=10)
        
        # 種別選択（チェックボックス）
        tk.Label(option_frame, text="出力する種別:", font=("Arial", 12)).pack(anchor=tk.W, padx=20)
        
        checkbox_frame = tk.Frame(option_frame)
        checkbox_frame.pack(pady=5)
        
        class_var = tk.BooleanVar(value=True)
        meeting_var = tk.BooleanVar(value=True)
        
        tk.Checkbutton(checkbox_frame, text="授業用", variable=class_var, 
                      font=("Arial", 11)).pack(side=tk.LEFT, padx=10)
        tk.Checkbutton(checkbox_frame, text="会議用", variable=meeting_var, 
                      font=("Arial", 11)).pack(side=tk.LEFT, padx=10)
        
        # エクスポートボタン
        export_btn = tk.Button(self.root, text="エクスポート", command=lambda: export_csv(),
                              font=("Arial", 14), bg="green", fg="white", width=15, height=2)
        export_btn.pack(pady=10)
        
        # 結果表示エリア
        result_text = tk.Text(self.root, height=12, width=70, font=("Arial", 10))
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
            
            # どちらか選択されているか確認
            export_class = class_var.get()
            export_meeting = meeting_var.get()
            
            if not export_class and not export_meeting:
                messagebox.showerror("エラー", "授業用または会議用のいずれかを選択してください")
                return
            
            # エクスポート実行
            result = ""
            
            if export_class:
                result += self.export_records_to_csv(date_str, "time_records")
                result += "\n" + "="*50 + "\n\n"
            
            if export_meeting:
                result += self.export_records_to_csv(date_str, "meeting_records")
            
            # 結果表示
            result_text.delete(1.0, tk.END)
            result_text.insert(1.0, result)
        
        # 戻るボタン
        back_btn = tk.Button(self.root, text="戻る", command=self.show_menu,
                           font=("Arial", 12))
        back_btn.pack(pady=10)
    
    def show_monthly_summary(self):
        """月次集計画面"""
        # 既存のウィジェットをクリア
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # タイトル
        title_label = tk.Label(self.root, text="月次集計", font=("Arial", 18, "bold"))
        title_label.pack(pady=20)
        
        # 入力フレーム
        input_frame = tk.Frame(self.root)
        input_frame.pack(pady=20)
        
        # 対象月入力
        tk.Label(input_frame, text="対象月 (YYYY-MM):", font=("Arial", 12)).pack(side=tk.LEFT, padx=5)
        month_entry = tk.Entry(input_frame, width=15, font=("Arial", 12))
        month_entry.pack(side=tk.LEFT, padx=5)
        
        # 今月をデフォルト表示
        current_month = datetime.now(JST).strftime("%Y-%m")
        month_entry.insert(0, current_month)
        
        # オプションフレーム
        option_frame = tk.Frame(self.root)
        option_frame.pack(pady=10)
        
        # 種別選択（チェックボックス）
        tk.Label(option_frame, text="出力する種別:", font=("Arial", 12)).pack(anchor=tk.W, padx=20)
        
        checkbox_frame = tk.Frame(option_frame)
        checkbox_frame.pack(pady=5)
        
        class_var = tk.BooleanVar(value=True)
        meeting_var = tk.BooleanVar(value=True)
        
        tk.Checkbutton(checkbox_frame, text="授業用", variable=class_var, 
                      font=("Arial", 11)).pack(side=tk.LEFT, padx=10)
        tk.Checkbutton(checkbox_frame, text="会議用", variable=meeting_var, 
                      font=("Arial", 11)).pack(side=tk.LEFT, padx=10)
        
        # 日次集計オプション
        daily_export_var = tk.BooleanVar(value=True)
        tk.Checkbutton(option_frame, text="日次集計も実行する", 
                      variable=daily_export_var, font=("Arial", 11)).pack(pady=5)
        
        # エクスポートボタン
        export_btn = tk.Button(self.root, text="エクスポート", command=lambda: export_monthly(),
                              font=("Arial", 14), bg="green", fg="white", width=15, height=2)
        export_btn.pack(pady=10)
        
        # 結果表示エリア
        result_text = tk.Text(self.root, height=12, width=70, font=("Arial", 10))
        result_text.pack(pady=10, padx=20)
        
        def export_monthly():
            """月次集計CSVエクスポート実行"""
            month_str = month_entry.get().strip()
            if not month_str:
                month_str = current_month
            
            # 月形式検証
            try:
                datetime.strptime(month_str, "%Y-%m")
            except ValueError:
                messagebox.showerror("エラー", "月形式が正しくありません (YYYY-MM)")
                return
            
            # どちらか選択されているか確認
            export_class = class_var.get()
            export_meeting = meeting_var.get()
            
            if not export_class and not export_meeting:
                messagebox.showerror("エラー", "授業用または会議用のいずれかを選択してください")
                return
            
            # エクスポート実行
            include_daily = daily_export_var.get()
            result = ""
            has_error = False
            
            if export_class:
                class_result = self.export_monthly_summary_to_csv(month_str, "time_records", include_daily)
                result += class_result
                result += "\n" + "="*50 + "\n\n"
                if "エラー" in class_result:
                    has_error = True
            
            if export_meeting:
                meeting_result = self.export_monthly_summary_to_csv(month_str, "meeting_records", include_daily)
                result += meeting_result
                result += "\n" + "="*50 + "\n\n"
                if "エラー" in meeting_result:
                    has_error = True
            
            # 両方選択されている場合は統合CSVも作成
            if export_class and export_meeting:
                summary_result = self.export_combined_monthly_summary(month_str)
                result += summary_result
                if "エラー" in summary_result:
                    has_error = True
            
            # 結果表示
            result_text.delete(1.0, tk.END)
            result_text.insert(1.0, result)
            result_text.see(tk.END)  # 最下部までスクロール
            
            # エラーがあった場合は警告表示
            if has_error:
                messagebox.showerror("エラー", 
                    "エクスポート中にエラーが発生しました。\n\n"
                    "・対象のCSVファイルがExcelなどで開かれている可能性があります\n"
                    "・ファイルを閉じてから再度実行してください\n\n"
                    "詳細は結果表示エリアを確認してください")
                return
            
            # エクスポート完了後の処理
            monthly_folder = os.path.join("monthly", month_str)
            if os.path.exists(monthly_folder):
                # システム終了確認
                if messagebox.askyesno("確認", "エクスポートが完了しました。\n\n出力フォルダを開いてシステムを終了しますか？"):
                    try:
                        os.startfile(os.path.abspath(monthly_folder))
                    except Exception as e:
                        print(f"フォルダを開くエラー: {e}")
                    finally:
                        self.root.quit()
        
        # 戻るボタン
        back_btn = tk.Button(self.root, text="戻る", command=self.show_menu,
                           font=("Arial", 12))
        back_btn.pack(pady=10)
    
    def toggle_sound_setting(self):
        """音量設定"""
        self.sound_enabled = not self.sound_enabled
        status = "有効" if self.sound_enabled else "無効"
        messagebox.showinfo("音量設定", f"音声機能を{status}にしました")
        
        if self.sound_enabled:
            self.play_beep("success")
        
        self.show_menu()
    
    def exit_app(self):
        """アプリケーション終了"""
        if messagebox.askyesno("確認", "アプリケーションを終了しますか？"):
            self.monitoring = False
            self.root.quit()
    
    # ===== データベース関連メソッド =====
    
    def init_database(self):
        """データベース初期化"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # time_records テーブル（授業用）
            cursor.execute("PRAGMA table_info(time_records)")
            columns = cursor.fetchall()
            column_names = [col[1] for col in columns]
            
            if not columns:
                cursor.execute('''
                    CREATE TABLE time_records (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        instructor_id INTEGER,
                        card_uid TEXT NOT NULL,
                        instructor_name TEXT NOT NULL,
                        record_type TEXT NOT NULL CHECK (record_type IN ('IN', 'OUT')),
                        timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                ''')
            else:
                if 'instructor_name' not in column_names:
                    cursor.execute("ALTER TABLE time_records ADD COLUMN instructor_name TEXT DEFAULT '未登録'")
                if 'instructor_id' not in column_names:
                    cursor.execute("ALTER TABLE time_records ADD COLUMN instructor_id INTEGER")
            
            # meeting_records テーブル（会議用）
            cursor.execute("PRAGMA table_info(meeting_records)")
            meeting_columns = cursor.fetchall()
            
            if not meeting_columns:
                cursor.execute('''
                    CREATE TABLE meeting_records (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        instructor_id INTEGER,
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
    
    def get_last_record(self, card_uid, table_name="time_records"):
        """最後の打刻記録を取得"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            query = f"SELECT record_type, timestamp FROM {table_name} WHERE card_uid = ? ORDER BY timestamp DESC LIMIT 1"
            cursor.execute(query, (card_uid,))
            
            result = cursor.fetchone()
            conn.close()
            
            if result:
                return {"type": result[0], "timestamp": result[1]}
            return None
            
        except Exception as e:
            print(f"最後の記録取得エラー: {e}")
            return None
    
    def record_attendance_to_db(self, card_uid, name, instructor_id, record_type, timestamp, table_name="time_records"):
        """データベースに打刻記録"""
        max_retries = 3
        for attempt in range(max_retries):
            try:
                conn = sqlite3.connect(self.db_path, timeout=10.0)
                cursor = conn.cursor()
                
                query = f"INSERT INTO {table_name} (instructor_id, card_uid, instructor_name, record_type, timestamp) VALUES (?, ?, ?, ?, ?)"
                cursor.execute(query, (instructor_id, card_uid, name, record_type, timestamp))
                
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
    
    def get_date_records(self, date_str, table_name="time_records"):
        """特定日付の打刻記録取得"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            query = f'''
                SELECT instructor_name, record_type, timestamp 
                FROM {table_name}
                WHERE DATE(timestamp) = ?
                ORDER BY timestamp DESC
            '''
            cursor.execute(query, (date_str,))
            
            results = cursor.fetchall()
            conn.close()
            return results
            
        except Exception as e:
            print(f"記録取得エラー: {e}")
            return []
    
    def get_date_summary(self, date_str, table_name="time_records"):
        """特定日付のサマリー取得"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            query = f'''
                SELECT instructor_name, record_type, timestamp
                FROM {table_name}
                WHERE DATE(timestamp) = ?
                ORDER BY instructor_name, timestamp
            '''
            cursor.execute(query, (date_str,))
            
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
    
    def export_records_to_csv(self, date_str, table_name="time_records"):
        """CSVエクスポート"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            query = f'''
                SELECT instructor_name, record_type, timestamp, card_uid
                FROM {table_name}
                WHERE DATE(timestamp) = ?
                ORDER BY timestamp
            '''
            cursor.execute(query, (date_str,))
            
            results = cursor.fetchall()
            conn.close()
            
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
                
                for name, record_type, timestamp, card_uid in results:
                    record_type_jp = "出勤" if record_type == "IN" else "退勤"
                    writer.writerow([name, record_type_jp, timestamp, card_uid])
            
            # 統計情報
            instructor_count = len(set(name for name, _, _, _ in results))
            in_count = sum(1 for _, record_type, _, _ in results if record_type == "IN")
            out_count = sum(1 for _, record_type, _, _ in results if record_type == "OUT")
            
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
    
    def export_monthly_summary_to_csv(self, month_str, table_name="time_records", include_daily=True):
        """月次集計CSVエクスポート（オプションで日次集計も実行可能）"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 対象月の日付一覧を取得
            query = f'''
                SELECT DISTINCT DATE(timestamp) as date
                FROM {table_name}
                WHERE strftime('%Y-%m', timestamp) = ?
                ORDER BY date
            '''
            cursor.execute(query, (month_str,))
            dates = [row[0] for row in cursor.fetchall()]
            
            if not dates:
                conn.close()
                return f"{month_str} の打刻記録はありません。"
            
            # ステップ1: 各日の日次集計を実行（オプション）
            result = f"=== 月次集計処理開始 ===\n\n"
            result += f"対象月: {month_str}\n"
            result += f"打刻記録がある日数: {len(dates)}日\n\n"
            
            daily_export_count = 0
            if include_daily:
                result += f"--- 日次集計処理 ---\n"
                
                for date in dates:
                    # 各日の日次集計を実行
                    daily_result = self.export_records_to_csv(date, table_name)
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
            
            query = f'''
                SELECT instructor_name, DATE(timestamp) as date
                FROM {table_name}
                WHERE strftime('%Y-%m', timestamp) = ?
                GROUP BY instructor_name, DATE(timestamp)
                ORDER BY instructor_name
            '''
            cursor.execute(query, (month_str,))
            
            results = cursor.fetchall()
            conn.close()
            
            # 講師ごとに出勤日数をカウント
            instructor_attendance = {}
            for name, date in results:
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
            if os.path.exists(csv_filename):
                file_mtime = os.path.getmtime(csv_filename)
                file_datetime = datetime.fromtimestamp(file_mtime)
                timestamp_str = file_datetime.strftime("%H%M%S")
                
                base_name = f"【{table_type_prefix}】出退勤記録_{month_str}"
                old_filename = os.path.join(old_dir, f"{base_name}_{timestamp_str}.csv")
                
                if os.path.exists(old_filename):
                    counter = 2
                    while True:
                        old_filename = os.path.join(old_dir, f"{base_name}_{timestamp_str}_{counter}.csv")
                        if not os.path.exists(old_filename):
                            break
                        counter += 1
                
                shutil.move(csv_filename, old_filename)
            
            # すべての講師を取得してID順にソート
            all_instructors = self.load_instructors_full()
            all_instructors_sorted = sorted(all_instructors, key=lambda x: int(x['instructor_id']))
            
            # まとめCSVファイルに書き込み（全講師の日次データを連結）
            import calendar
            year, month = map(int, month_str.split('-'))
            _, last_day = calendar.monthrange(year, month)
            
            # 各講師の日次データを取得
            conn_daily = sqlite3.connect(self.db_path)
            cursor_daily = conn_daily.cursor()
            
            with open(csv_filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.writer(csvfile)
                
                # ヘッダー行を書き込み
                writer.writerow(['講師ID', '講師名', '日付', '出社時刻', '退社時刻', '外出時刻', '復帰時刻', '備考'])
                
                # すべての講師について日次データを出力
                for instructor in all_instructors_sorted:
                    instructor_id = instructor['instructor_id']
                    instructor_name = instructor['name']
                    
                    # 指定されたテーブルから打刻時刻を取得
                    query_daily = f'''
                        SELECT DATE(timestamp) as date, TIME(timestamp) as time
                        FROM {table_name}
                        WHERE instructor_id = ? AND strftime('%Y-%m', timestamp) = ?
                        ORDER BY date, time
                    '''
                    cursor_daily.execute(query_daily, (instructor_id, month_str))
                    records_daily = cursor_daily.fetchall()
                    
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
                        
                        # その日の打刻時刻を取得
                        if date_str in daily_data_per_instructor:
                            times = sorted(daily_data_per_instructor[date_str])
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
            
            conn_daily.close()
            
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
            
            # 各講師の日次データを集計
            instructor_daily_count = 0
            for instructor in all_instructors_sorted:
                instructor_id = instructor['instructor_id']
                instructor_name = instructor['name']
                
                # 講師別日次集計ファイルを作成（授業用 or 会議用）
                individual_result = self.export_instructor_daily_summary(
                    month_str, instructor_id, instructor_name, month_subdir, table_name
                )
                if individual_result:
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
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 授業記録を取得（1日1回まで）
            class_query = '''
                SELECT instructor_id, instructor_name, DATE(timestamp) as date
                FROM time_records
                WHERE strftime('%Y-%m', timestamp) = ?
                GROUP BY instructor_id, instructor_name, DATE(timestamp)
            '''
            cursor.execute(class_query, (month_str,))
            class_results = cursor.fetchall()
            
            # 会議記録を取得（1日1回まで）
            meeting_query = '''
                SELECT instructor_id, instructor_name, DATE(timestamp) as date
                FROM meeting_records
                WHERE strftime('%Y-%m', timestamp) = ?
                GROUP BY instructor_id, instructor_name, DATE(timestamp)
            '''
            cursor.execute(meeting_query, (month_str,))
            meeting_results = cursor.fetchall()
            
            conn.close()
            
            # 講師ごとに集計
            instructor_summary = {}
            
            # 授業回数を集計（日付ごとに1回のみカウント）
            for instructor_id, name, date in class_results:
                key = str(instructor_id)
                if key not in instructor_summary:
                    instructor_summary[key] = {'class_dates': set(), 'meeting_dates': set()}
                instructor_summary[key]['class_dates'].add(date)
            
            # 会議回数を集計（日付ごとに1回のみカウント）
            for instructor_id, name, date in meeting_results:
                key = str(instructor_id)
                if key not in instructor_summary:
                    instructor_summary[key] = {'class_dates': set(), 'meeting_dates': set()}
                instructor_summary[key]['meeting_dates'].add(date)
            
            # すべての講師を取得してID順にソート
            all_instructors = self.load_instructors_full()
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
            if os.path.exists(csv_filename):
                file_mtime = os.path.getmtime(csv_filename)
                file_datetime = datetime.fromtimestamp(file_mtime)
                timestamp_str = file_datetime.strftime("%H%M%S")
                
                base_name = f"【まとめ】出退勤記録_{month_str}"
                old_filename = os.path.join(old_dir, f"{base_name}_{timestamp_str}.csv")
                
                if os.path.exists(old_filename):
                    counter = 2
                    while True:
                        old_filename = os.path.join(old_dir, f"{base_name}_{timestamp_str}_{counter}.csv")
                        if not os.path.exists(old_filename):
                            break
                        counter += 1
                
                shutil.move(csv_filename, old_filename)
            
            # CSVファイルに書き込み
            with open(csv_filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(['【まとめ】講師ID', '【まとめ】講師名', '授業回数', '時間数（3.25h）', '会議回数', '出勤回数'])
                
                # すべての講師について出力（出勤記録がない講師も含む）
                for instructor in all_instructors_sorted:
                    instructor_id = instructor['instructor_id']
                    name = instructor['name']
                    
                    summary = instructor_summary.get(instructor_id, {'class_dates': set(), 'meeting_dates': set()})
                    class_dates = summary['class_dates']
                    meeting_dates = summary['meeting_dates']
                    
                    class_count = len(class_dates)
                    meeting_count = len(meeting_dates)
                    # 出勤回数は授業日と会議日の和集合（1日1回まで）
                    total_attendance = len(class_dates | meeting_dates)
                    hours = class_count * 3.25
                    
                    writer.writerow([instructor_id, name, class_count, hours, meeting_count, total_attendance])
            
            # 統計情報
            total_instructors_registered = len(all_instructors_sorted)
            total_class_days = sum(summary.get('class_count', 0) for summary in instructor_summary.values())
            total_meeting_days = sum(summary.get('meeting_count', 0) for summary in instructor_summary.values())
            
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
    
    def export_instructor_daily_summary(self, month_str, instructor_id, instructor_name, output_dir, table_name="time_records"):
        """講師別日次集計CSVエクスポート（授業用または会議用）"""
        try:
            import calendar
            
            # 月の日数を取得
            year, month = map(int, month_str.split('-'))
            _, last_day = calendar.monthrange(year, month)
            
            # データベースから打刻記録を取得（授業用 or 会議用）
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 指定されたテーブルから打刻時刻を取得
            query = f'''
                SELECT DATE(timestamp) as date, TIME(timestamp) as time
                FROM {table_name}
                WHERE instructor_id = ? AND strftime('%Y-%m', timestamp) = ?
                ORDER BY date, time
            '''
            cursor.execute(query, (instructor_id, month_str))
            records = cursor.fetchall()
            
            conn.close()
            
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
    
    # ===== カードリーダー関連メソッド =====
    
    def connect_to_card(self, reader):
        """カードに接続"""
        try:
            connection = reader.createConnection()
            connection.connect()
            return connection
        except (NoCardException, CardConnectionException):
            return None
        except Exception:
            return None
    
    def get_card_uid(self, connection):
        """カードUID取得"""
        if not connection:
            return None
            
        try:
            get_uid = [0xFF, 0xCA, 0x00, 0x00, 0x00]
            data, sw1, sw2 = connection.transmit(get_uid)
            
            if sw1 == 0x90 and sw2 == 0x00:
                uid = toHexString(data)
                return uid
            else:
                return None
                
        except Exception:
            return None
    
    def disconnect(self, connection):
        """接続切断"""
        if connection:
            try:
                connection.disconnect()
            except:
                pass
    
    def is_card_present(self, reader):
        """カード存在チェック"""
        try:
            test_connection = reader.createConnection()
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