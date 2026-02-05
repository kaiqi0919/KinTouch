# 出退勤管理システム（GUI版・複数リーダー対応）
# モジュール化バージョン
# pip install pyscard

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import time
import os
import threading
import hashlib
from datetime import datetime

# モジュールのインポート
from modules import (
    JST, PASSWORD_HASH, DATA_DIR, CONFIG_PATH, WINDOW_WIDTH, WINDOW_HEIGHT,
    DatabaseManager, CardReaderManager, CSVExporter, MonthlyExporter,
    ConfigManager, SoundManager
)

class AttendanceSystemGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("出退勤確認システム（複数リーダー対応）")
        
        # ウィンドウを画面中央に配置
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        center_x = int((screen_width - WINDOW_WIDTH) / 2)
        center_y = int((screen_height - WINDOW_HEIGHT) / 2)
        self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+{center_x}+{center_y}")
        
        # ディレクトリ初期化
        if not os.path.exists(DATA_DIR):
            os.makedirs(DATA_DIR)
        
        # マネージャー初期化
        self.db_manager = DatabaseManager(os.path.join(DATA_DIR, "attendance.db"))
        self.card_reader_manager = CardReaderManager()
        self.csv_exporter = CSVExporter(self.db_manager)
        self.monthly_exporter = MonthlyExporter(self.db_manager, self.csv_exporter)
        self.config_manager = ConfigManager(CONFIG_PATH)
        self.sound_manager = SoundManager()
        
        # 監視関連
        self.monitoring = False
        self.clear_timer_class = None
        self.clear_timer_meeting = None
        
        # 設定の読み込みと初期化
        config = self.config_manager.load_config()
        if not config:
            self.show_reader_setup()
        else:
            if not self.card_reader_manager.initialize_readers(
                config['class_reader'], config['meeting_reader']
            ):
                messagebox.showwarning("警告", "リーダーの初期化に失敗しました\n設定を確認してください")
                self.show_reader_setup()
            else:
                self.show_menu()
    
    def show_reader_setup(self):
        """リーダー設定画面"""
        for widget in self.root.winfo_children():
            widget.destroy()
        
        tk.Label(self.root, text="リーダー設定", font=("Arial", 18, "bold")).pack(pady=20)
        tk.Label(self.root, text="接続されているリーダーを授業用と会議用に割り当ててください",
                 font=("Arial", 12)).pack(pady=10)
        
        try:
            reader_names = self.card_reader_manager.get_available_readers()
            if len(reader_names) < 2:
                tk.Label(self.root, 
                    text=f"エラー: {len(reader_names)}台のリーダーしか検出されていません\n2台のリーダーを接続してください",
                    font=("Arial", 12), fg="red").pack(pady=20)
                tk.Button(self.root, text="再試行", 
                    command=self.show_reader_setup, font=("Arial", 12)).pack(pady=10)
                return
            
            select_frame = tk.Frame(self.root)
            select_frame.pack(pady=20)
            
            tk.Label(select_frame, text="授業用リーダー:", 
                    font=("Arial", 12)).grid(row=0, column=0, padx=10, pady=10, sticky='e')
            class_combo = ttk.Combobox(select_frame, values=reader_names, 
                                      width=40, font=("Arial", 10), state='readonly')
            class_combo.grid(row=0, column=1, padx=10, pady=10)
            if reader_names:
                class_combo.set(reader_names[0])
            
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
                
                if self.config_manager.save_config(class_reader, meeting_reader):
                    if self.card_reader_manager.initialize_readers(class_reader, meeting_reader):
                        messagebox.showinfo("成功", "リーダー設定が完了しました")
                        self.show_menu()
                    else:
                        messagebox.showerror("エラー", "リーダーの初期化に失敗しました")
                else:
                    messagebox.showerror("エラー", "設定の保存に失敗しました")
            
            tk.Button(self.root, text="設定を保存", command=save_and_continue,
                     font=("Arial", 14), bg="green", fg="white", width=15).pack(pady=20)
            
        except Exception as e:
            tk.Label(self.root, 
                text=f"エラー: {e}\nリーダーの接続を確認してください",
                font=("Arial", 12), fg="red").pack(pady=20)
    
    def show_menu(self):
        """メニュー画面を表示"""
        for widget in self.root.winfo_children():
            widget.destroy()
        
        tk.Label(self.root, text="出退勤確認システム", font=("Arial", 20, "bold")).pack(pady=20)
        
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)
        
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
            tk.Button(button_frame, text=text, width=20, height=2,
                     font=("Arial", 12), command=command).grid(row=row, column=col, padx=10, pady=10)
        
        sound_status = "有効" if self.sound_manager.sound_enabled else "無効"
        tk.Label(self.root, text=f"音声: {sound_status}", font=("Arial", 10)).pack(side=tk.BOTTOM, pady=10)
    
    def show_attendance_monitor(self):
        """打刻受付画面（2分割）"""
        for widget in self.root.winfo_children():
            widget.destroy()
        
        tk.Label(self.root, text="打刻受付", font=("Arial", 18, "bold")).pack(pady=10)
        
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
        
        tk.Button(self.root, text="終了", command=self.stop_monitoring,
                 font=("Arial", 12), bg="red", fg="white").pack(pady=10)
        
        # 監視開始
        self.monitoring = True
        threading.Thread(
            target=self.monitor_cards, 
            args=(self.card_reader_manager.class_reader, class_status, 'class'), 
            daemon=True
        ).start()
        threading.Thread(
            target=self.monitor_cards, 
            args=(self.card_reader_manager.meeting_reader, meeting_status, 'meeting'), 
            daemon=True
        ).start()
    
    def monitor_cards(self, reader, status_label, reader_type):
        """カード監視スレッド"""
        connection = None
        last_uid = None
        
        while self.monitoring:
            try:
                if self.card_reader_manager.is_card_present(reader):
                    connection = self.card_reader_manager.connect_to_card(reader)
                    if connection:
                        uid = self.card_reader_manager.get_card_uid(connection)
                        if uid and uid != last_uid:
                            self.sound_manager.play_beep("card_detected")
                            self.process_attendance(uid, status_label, reader_type)
                            last_uid = uid
                        self.card_reader_manager.disconnect(connection)
                        connection = None
                else:
                    if last_uid:
                        last_uid = None
                
                time.sleep(0.5)
            except Exception as e:
                print(f"監視エラー ({reader_type}): {e}")
                time.sleep(1)
        
        if connection:
            self.card_reader_manager.disconnect(connection)
    
    def process_attendance(self, uid, status_label, reader_type):
        """打刻処理"""
        instructor_info = self.db_manager.get_instructor_info_by_uid(uid)
        
        if not instructor_info:
            self.root.after(0, lambda: status_label.config(
                text="未登録のカードです", fg="red"))
            self.sound_manager.play_beep("error")
            self.root.after(2000, lambda: status_label.config(
                text="カードをかざしてください...", fg="blue"))
            return
        
        table_name = "time_records" if reader_type == "class" else "meeting_records"
        
        last_record = self.db_manager.get_last_record(uid, table_name)
        if last_record is None or last_record["type"] == "OUT":
            record_type = "IN"
            action = "出勤"
            action_color = "green"
        else:
            record_type = "OUT"
            action = "退勤"
            action_color = "orange"
        
        jst_now = datetime.now(JST)
        timestamp_str = jst_now.strftime("%Y-%m-%d %H:%M:%S")
        
        if self.db_manager.record_attendance_to_db(uid, instructor_info['name'], 
                                       instructor_info['instructor_id'], 
                                       record_type, timestamp_str, table_name):
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
            self.sound_manager.play_beep("success")
        else:
            self.sound_manager.play_beep("error")
    
    def display_attendance_info(self, instructor_id, name, uid, timestamp, action, color, status_label, reader_type):
        """打刻情報を3秒間表示"""
        if reader_type == "class":
            info_labels = self.class_info_labels
            timer_attr = 'clear_timer_class'
        else:
            info_labels = self.meeting_info_labels
            timer_attr = 'clear_timer_meeting'
        
        timer_id = getattr(self, timer_attr, None)
        if timer_id is not None:
            self.root.after_cancel(timer_id)
        
        status_label.config(text=f"{action}記録完了！", fg=color)
        
        info_labels['instructor_id'].config(text=f"講師番号: {instructor_id}")
        info_labels['name'].config(text=f"{name}")
        info_labels['timestamp'].config(text=f"{timestamp}")
        info_labels['action'].config(text=f"【{action}】", fg=color)
        
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
        time.sleep(0.5)
        self.show_menu()
    
    def show_instructor_list(self):
        """講師一覧画面"""
        for widget in self.root.winfo_children():
            widget.destroy()
        
        tk.Label(self.root, text="講師一覧", font=("Arial", 18, "bold")).pack(pady=10)
        
        register_btn = tk.Button(self.root, text="講師登録", 
                               command=self.show_instructor_registration,
                               font=("Arial", 12), bg="green", fg="white")
        register_btn.place(x=680, y=10)
        
        table_frame = tk.Frame(self.root)
        table_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(table_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        columns = ('講師番号', 'カードUID', '講師名', '登録日時')
        tree = ttk.Treeview(table_frame, columns=columns, show='headings',
                           yscrollcommand=scrollbar.set)
        
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=180)
        
        instructors = self.db_manager.load_instructors_full()
        for instructor in instructors:
            tree.insert('', tk.END, values=(
                instructor['instructor_id'],
                instructor['card_uid'],
                instructor['name'],
                instructor['created_at']
            ))
        
        tree.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=tree.yview)
        
        tk.Label(self.root, text=f"登録講師数: {len(instructors)}人",
                 font=("Arial", 12)).pack(pady=5)
        
        tk.Button(self.root, text="戻る", command=self.show_menu,
                 font=("Arial", 12)).pack(pady=10)
    
    def show_instructor_registration(self):
        """講師登録画面"""
        reg_window = tk.Toplevel(self.root)
        reg_window.title("講師登録")
        reg_window.geometry("450x350")
        
        tk.Label(reg_window, text="講師登録", font=("Arial", 16, "bold")).pack(pady=10)
        
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
        
        input_frame = tk.Frame(reg_window)
        input_frame.pack(pady=10)
        
        tk.Label(input_frame, text="カードUID:").grid(row=0, column=0, padx=5, pady=5)
        uid_entry = tk.Entry(input_frame, width=30, state='readonly')
        uid_entry.grid(row=0, column=1, padx=5, pady=5)
        
        next_id = self.db_manager.get_next_instructor_id()
        tk.Label(input_frame, text="講師番号:").grid(row=1, column=0, padx=5, pady=5)
        id_entry = tk.Entry(input_frame, width=30)
        id_entry.insert(0, str(next_id))
        id_entry.grid(row=1, column=1, padx=5, pady=5)
        
        tk.Label(input_frame, text="講師名:").grid(row=2, column=0, padx=5, pady=5)
        name_entry = tk.Entry(input_frame, width=30)
        name_entry.grid(row=2, column=1, padx=5, pady=5)
        
        detected_uid = {'uid': None}
        
        def check_card():
            if reg_window.winfo_exists():
                selected_reader = self.card_reader_manager.class_reader if reader_var.get() == "class" else self.card_reader_manager.meeting_reader
                
                if self.card_reader_manager.is_card_present(selected_reader):
                    connection = self.card_reader_manager.connect_to_card(selected_reader)
                    if connection:
                        uid = self.card_reader_manager.get_card_uid(connection)
                        if uid and uid != detected_uid['uid']:
                            detected_uid['uid'] = uid
                            uid_entry.config(state='normal')
                            uid_entry.delete(0, tk.END)
                            uid_entry.insert(0, uid)
                            uid_entry.config(state='readonly')
                            status_label.config(text="カード検出！講師情報を入力してください", fg="green")
                            self.sound_manager.play_beep("card_detected")
                        self.card_reader_manager.disconnect(connection)
                reg_window.after(500, check_card)
        
        check_card()
        
        def register():
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
            
            if self.db_manager.add_instructor_with_id(instructor_id, uid, name):
                messagebox.showinfo("成功", "講師を登録しました")
                reg_window.destroy()
                self.show_instructor_list()
            else:
                messagebox.showerror("エラー", "登録に失敗しました")
        
        btn_frame = tk.Frame(reg_window)
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="登録", command=register, 
                 font=("Arial", 12), bg="green", fg="white", width=10).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="キャンセル", command=reg_window.destroy,
                 font=("Arial", 12), width=10).pack(side=tk.LEFT, padx=5)
    
    def show_attendance_records(self):
        """打刻表示画面（2分割）"""
        for widget in self.root.winfo_children():
            widget.destroy()
        
        tk.Label(self.root, text="打刻表示", font=("Arial", 18, "bold")).pack(pady=10)
        
        input_frame = tk.Frame(self.root)
        input_frame.pack(pady=10)
        
        tk.Label(input_frame, text="日付 (YYYY-MM-DD):", font=("Arial", 12)).pack(side=tk.LEFT, padx=5)
        date_entry = tk.Entry(input_frame, width=15, font=("Arial", 12))
        date_entry.pack(side=tk.LEFT, padx=5)
        
        today = datetime.now(JST).strftime("%Y-%m-%d")
        date_entry.insert(0, today)
        
        main_frame = tk.Frame(self.root)
        main_frame.pack(pady=10, fill=tk.BOTH, expand=True, padx=10)
        
        # 左側：授業用
        class_frame = tk.Frame(main_frame, relief=tk.RIDGE, borderwidth=2)
        class_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        tk.Label(class_frame, text="授業用", font=("Arial", 14, "bold"), 
                bg="lightblue").pack(fill=tk.X, pady=5)
        
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
            date_str = date_entry.get().strip()
            if not date_str:
                date_str = today
            
            try:
                datetime.strptime(date_str, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("エラー", "日付形式が正しくありません (YYYY-MM-DD)")
                return
            
            class_records = self.db_manager.get_date_records(date_str, "time_records")
            
            for item in class_tree.get_children():
                class_tree.delete(item)
            
            for name, record_type, timestamp in class_records:
                action = "出勤" if record_type == "IN" else "退勤"
                time_part = timestamp.split()[1] if ' ' in timestamp else timestamp
                class_tree.insert('', tk.END, values=(time_part, name, action))
            
            class_count_label.config(text=f"記録数: {len(class_records)}件")
            
            meeting_records = self.db_manager.get_date_records(date_str, "meeting_records")
            
            for item in meeting_tree.get_children():
                meeting_tree.delete(item)
            
            for name, record_type, timestamp in meeting_records:
                action = "出勤" if record_type == "IN" else "退勤"
                time_part = timestamp.split()[1] if ' ' in timestamp else timestamp
                meeting_tree.insert('', tk.END, values=(time_part, name, action))
            
            meeting_count_label.config(text=f"記録数: {len(meeting_records)}件")
        
        tk.Button(input_frame, text="表示", command=display_records,
                 font=("Arial", 12), bg="blue", fg="white").pack(side=tk.LEFT, padx=5)
        
        display_records()
        
        tk.Button(self.root, text="戻る", command=self.show_menu,
                 font=("Arial", 12)).pack(pady=10)
    
    def show_attendance_summary(self):
        """打刻サマリー画面（2分割）"""
        for widget in self.root.winfo_children():
            widget.destroy()
        
        tk.Label(self.root, text="打刻サマリー", font=("Arial", 18, "bold")).pack(pady=10)
        
        input_frame = tk.Frame(self.root)
        input_frame.pack(pady=10)
        
        tk.Label(input_frame, text="日付 (YYYY-MM-DD):", font=("Arial", 12)).pack(side=tk.LEFT, padx=5)
        date_entry = tk.Entry(input_frame, width=15, font=("Arial", 12))
        date_entry.pack(side=tk.LEFT, padx=5)
        
        today = datetime.now(JST).strftime("%Y-%m-%d")
        date_entry.insert(0, today)
        
        main_frame = tk.Frame(self.root)
        main_frame.pack(pady=10, fill=tk.BOTH, expand=True, padx=10)
        
        # 左側：授業用
        class_frame = tk.Frame(main_frame, relief=tk.RIDGE, borderwidth=2)
        class_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        tk.Label(class_frame, text="授業用", font=("Arial", 14, "bold"), 
                bg="lightblue").pack(fill=tk.X, pady=5)
        
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
            date_str = date_entry.get().strip()
            if not date_str:
                date_str = today
            
            try:
                datetime.strptime(date_str, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("エラー", "日付形式が正しくありません (YYYY-MM-DD)")
                return
            
            class_summary = self.db_manager.get_date_summary(date_str, "time_records")
            
            for item in class_tree.get_children():
                class_tree.delete(item)
            
            for name, status, last_time, record_str in class_summary:
                class_tree.insert('', tk.END, values=(name, status, last_time, record_str))
            
            class_count_label.config(text=f"打刻した講師数: {len(class_summary)}人")
            
            meeting_summary = self.db_manager.get_date_summary(date_str, "meeting_records")
            
            for item in meeting_tree.get_children():
                meeting_tree.delete(item)
            
            for name, status, last_time, record_str in meeting_summary:
                meeting_tree.insert('', tk.END, values=(name, status, last_time, record_str))
            
            meeting_count_label.config(text=f"打刻した講師数: {len(meeting_summary)}人")
        
        tk.Button(input_frame, text="表示", command=display_summary,
                 font=("Arial", 12), bg="blue", fg="white").pack(side=tk.LEFT, padx=5)
        
        display_summary()
        
        tk.Button(self.root, text="戻る", command=self.show_menu,
                 font=("Arial", 12)).pack(pady=10)
    
    def show_attendance_correction(self):
        """打刻修正画面（パスワード認証付き）"""
        password = simpledialog.askstring("パスワード入力", "パスワードを入力してください:", show='*')
        
        if password is None:
            return
        
        password_hash = hashlib.sha256(password.encode('utf-8')).hexdigest()
        if password_hash != PASSWORD_HASH:
            messagebox.showerror("エラー", "パスワードが正しくありません")
            return
        
        for widget in self.root.winfo_children():
            widget.destroy()
        
        tk.Label(self.root, text="打刻修正", font=("Arial", 18, "bold")).pack(pady=20)
        
        input_frame = tk.Frame(self.root)
        input_frame.pack(pady=20)
        
        tk.Label(input_frame, text="種別:", font=("Arial", 12)).grid(row=0, column=0, padx=10, pady=5, sticky='e')
        table_var = tk.StringVar(value="time_records")
        type_frame = tk.Frame(input_frame)
        type_frame.grid(row=0, column=1, padx=10, pady=5, sticky='w')
        tk.Radiobutton(type_frame, text="授業用", variable=table_var, 
                      value="time_records", font=("Arial", 11)).pack(side=tk.LEFT)
        tk.Radiobutton(type_frame, text="会議用", variable=table_var, 
                      value="meeting_records", font=("Arial", 11)).pack(side=tk.LEFT)
        
        tk.Label(input_frame, text="講師選択 (必須):", font=("Arial", 12)).grid(row=1, column=0, padx=10, pady=10, sticky='e')
        
        instructors = self.db_manager.load_instructors_full()
        instructor_options = [f"{inst['instructor_id']}: {inst['name']}" for inst in instructors]
        
        instructor_combo = ttk.Combobox(input_frame, values=instructor_options, width=28, font=("Arial", 11), state='readonly')
        instructor_combo.grid(row=1, column=1, padx=10, pady=10)
        if instructor_options:
            instructor_combo.set("講師を選択してください")
        
        tk.Label(input_frame, text="時刻 (YYYY-MM-DD HH:MM:SS):", font=("Arial", 12)).grid(row=2, column=0, padx=10, pady=10, sticky='e')
        tk.Label(input_frame, text="※空欄の場合は現在時刻", font=("Arial", 9), fg="gray").grid(row=3, column=1, sticky='w')
        time_entry = tk.Entry(input_frame, width=30, font=("Arial", 12))
        time_entry.grid(row=2, column=1, padx=10, pady=10)
        
        tk.Label(input_frame, text="打刻種別:", font=("Arial", 12)).grid(row=4, column=0, padx=10, pady=10, sticky='e')
        record_type_var = tk.StringVar(value="IN")
        tk.Radiobutton(input_frame, text="出勤", variable=record_type_var, value="IN", font=("Arial", 11)).grid(row=4, column=1, sticky='w')
        tk.Radiobutton(input_frame, text="退勤", variable=record_type_var, value="OUT", font=("Arial", 11)).grid(row=5, column=1, sticky='w')
        
        def register_correction():
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
            
            instructor_info = self.db_manager.get_instructor_info_by_id(instructor_id)
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
            
            if self.db_manager.record_attendance_to_db(instructor_info['card_uid'], 
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
        
        tk.Button(self.root, text="登録", command=register_correction,
                 font=("Arial", 14), bg="green", fg="white", width=15, height=2).pack(pady=20)
        
        tk.Button(self.root, text="戻る", command=self.show_menu,
                 font=("Arial", 12)).pack(pady=10)
    
    def show_csv_export(self):
        """日次集計画面"""
        for widget in self.root.winfo_children():
            widget.destroy()
        
        tk.Label(self.root, text="日次集計", font=("Arial", 18, "bold")).pack(pady=20)
        
        input_frame = tk.Frame(self.root)
        input_frame.pack(pady=20)
        
        tk.Label(input_frame, text="対象日 (YYYY-MM-DD):", font=("Arial", 12)).pack(side=tk.LEFT, padx=5)
        date_entry = tk.Entry(input_frame, width=15, font=("Arial", 12))
        date_entry.pack(side=tk.LEFT, padx=5)
        
        today = datetime.now(JST).strftime("%Y-%m-%d")
        date_entry.insert(0, today)
        
        option_frame = tk.Frame(self.root)
        option_frame.pack(pady=10)
        
        tk.Label(option_frame, text="出力する種別:", font=("Arial", 12)).pack(anchor=tk.W, padx=20)
        
        checkbox_frame = tk.Frame(option_frame)
        checkbox_frame.pack(pady=5)
        
        class_var = tk.BooleanVar(value=True)
        meeting_var = tk.BooleanVar(value=True)
        
        tk.Checkbutton(checkbox_frame, text="授業用", variable=class_var, 
                      font=("Arial", 11)).pack(side=tk.LEFT, padx=10)
        tk.Checkbutton(checkbox_frame, text="会議用", variable=meeting_var, 
                      font=("Arial", 11)).pack(side=tk.LEFT, padx=10)
        
        export_btn = tk.Button(self.root, text="エクスポート", command=lambda: export_csv(),
                              font=("Arial", 14), bg="green", fg="white", width=15, height=2)
        export_btn.pack(pady=10)
        
        result_text = tk.Text(self.root, height=12, width=70, font=("Arial", 10))
        result_text.pack(pady=10, padx=20)
        
        def export_csv():
            date_str = date_entry.get().strip()
            if not date_str:
                date_str = today
            
            try:
                datetime.strptime(date_str, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("エラー", "日付形式が正しくありません (YYYY-MM-DD)")
                return
            
            export_class = class_var.get()
            export_meeting = meeting_var.get()
            
            if not export_class and not export_meeting:
                messagebox.showerror("エラー", "授業用または会議用のいずれかを選択してください")
                return
            
            result = ""
            
            if export_class:
                result += self.csv_exporter.export_records_to_csv(date_str, "time_records")
                result += "\n" + "="*50 + "\n\n"
            
            if export_meeting:
                result += self.csv_exporter.export_records_to_csv(date_str, "meeting_records")
            
            result_text.delete(1.0, tk.END)
            result_text.insert(1.0, result)
        
        tk.Button(self.root, text="戻る", command=self.show_menu,
                 font=("Arial", 12)).pack(pady=10)
    
    def show_monthly_summary(self):
        """月次集計画面"""
        for widget in self.root.winfo_children():
            widget.destroy()
        
        tk.Label(self.root, text="月次集計", font=("Arial", 18, "bold")).pack(pady=20)
        
        input_frame = tk.Frame(self.root)
        input_frame.pack(pady=20)
        
        tk.Label(input_frame, text="対象月 (YYYY-MM):", font=("Arial", 12)).pack(side=tk.LEFT, padx=5)
        month_entry = tk.Entry(input_frame, width=15, font=("Arial", 12))
        month_entry.pack(side=tk.LEFT, padx=5)
        
        current_month = datetime.now(JST).strftime("%Y-%m")
        month_entry.insert(0, current_month)
        
        option_frame = tk.Frame(self.root)
        option_frame.pack(pady=10)
        
        tk.Label(option_frame, text="出力する種別:", font=("Arial", 12)).pack(anchor=tk.W, padx=20)
        
        checkbox_frame = tk.Frame(option_frame)
        checkbox_frame.pack(pady=5)
        
        class_var = tk.BooleanVar(value=True)
        meeting_var = tk.BooleanVar(value=True)
        
        tk.Checkbutton(checkbox_frame, text="授業用", variable=class_var, 
                      font=("Arial", 11)).pack(side=tk.LEFT, padx=10)
        tk.Checkbutton(checkbox_frame, text="会議用", variable=meeting_var, 
                      font=("Arial", 11)).pack(side=tk.LEFT, padx=10)
        
        daily_export_var = tk.BooleanVar(value=True)
        tk.Checkbutton(option_frame, text="日次集計も実行する", 
                      variable=daily_export_var, font=("Arial", 11)).pack(pady=5)
        
        export_btn = tk.Button(self.root, text="エクスポート", command=lambda: export_monthly(),
                              font=("Arial", 14), bg="green", fg="white", width=15, height=2)
        export_btn.pack(pady=10)
        
        result_text = tk.Text(self.root, height=12, width=70, font=("Arial", 10))
        result_text.pack(pady=10, padx=20)
        
        def export_monthly():
            month_str = month_entry.get().strip()
            if not month_str:
                month_str = current_month
            
            try:
                datetime.strptime(month_str, "%Y-%m")
            except ValueError:
                messagebox.showerror("エラー", "月形式が正しくありません (YYYY-MM)")
                return
            
            export_class = class_var.get()
            export_meeting = meeting_var.get()
            
            if not export_class and not export_meeting:
                messagebox.showerror("エラー", "授業用または会議用のいずれかを選択してください")
                return
            
            include_daily = daily_export_var.get()
            result = ""
            has_error = False
            
            if export_class:
                class_result = self.monthly_exporter.export_monthly_summary_to_csv(month_str, "time_records", include_daily)
                result += class_result
                result += "\n" + "="*50 + "\n\n"
                if "エラー" in class_result:
                    has_error = True
            
            if export_meeting:
                meeting_result = self.monthly_exporter.export_monthly_summary_to_csv(month_str, "meeting_records", include_daily)
                result += meeting_result
                result += "\n" + "="*50 + "\n\n"
                if "エラー" in meeting_result:
                    has_error = True
            
            if export_class and export_meeting:
                summary_result = self.monthly_exporter.export_combined_monthly_summary(month_str)
                result += summary_result
                if "エラー" in summary_result:
                    has_error = True
            
            result_text.delete(1.0, tk.END)
            result_text.insert(1.0, result)
            result_text.see(tk.END)
            
            if has_error:
                messagebox.showerror("エラー", 
                    "エクスポート中にエラーが発生しました。\n\n"
                    "・対象のCSVファイルがExcelなどで開かれている可能性があります\n"
                    "・ファイルを閉じてから再度実行してください\n\n"
                    "詳細は結果表示エリアを確認してください")
                return
            
            monthly_folder = os.path.join("monthly", month_str)
            if os.path.exists(monthly_folder):
                if messagebox.askyesno("確認", "エクスポートが完了しました。\n\n出力フォルダを開いてシステムを終了しますか？"):
                    try:
                        os.startfile(os.path.abspath(monthly_folder))
                    except Exception as e:
                        print(f"フォルダを開くエラー: {e}")
                    finally:
                        self.root.quit()
        
        tk.Button(self.root, text="戻る", command=self.show_menu,
                 font=("Arial", 12)).pack(pady=10)
    
    def toggle_sound_setting(self):
        """音量設定"""
        enabled = self.sound_manager.toggle_sound()
        status = "有効" if enabled else "無効"
        messagebox.showinfo("音量設定", f"音声機能を{status}にしました")
        
        if enabled:
            self.sound_manager.play_beep("success")
        
        self.show_menu()
    
    def exit_app(self):
        """アプリケーション終了"""
        if messagebox.askyesno("確認", "アプリケーションを終了しますか？"):
            self.monitoring = False
            self.root.quit()

def main():
    """メイン関数"""
    root = tk.Tk()
    app = AttendanceSystemGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()