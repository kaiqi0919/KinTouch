# 出退勤管理システム - 打刻修正管理モジュール

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import hashlib
from datetime import datetime
from modules.constants import JST, PASSWORD_HASH

class CorrectionManager:
    """打刻修正管理クラス"""
    
    def __init__(self, root, db_manager, card_reader_manager, sound_manager, show_menu_callback):
        self.root = root
        self.db_manager = db_manager
        self.card_reader_manager = card_reader_manager
        self.sound_manager = sound_manager
        self.show_menu_callback = show_menu_callback
    
    def show_attendance_correction(self):
        """打刻修正画面（パスワードまたはマスターキーカード認証）"""
        self.show_correction_auth()
    
    def show_correction_auth(self):
        """打刻修正の認証画面"""
        for widget in self.root.winfo_children():
            widget.destroy()
        
        tk.Label(self.root, text="打刻修正 - 認証", font=("Arial", 18, "bold")).pack(pady=20)
        
        tk.Label(self.root, text="パスワードを入力するか、マスターキーカードをかざしてください", 
                 font=("Arial", 12)).pack(pady=10)
        
        # パスワード入力フレーム
        password_frame = tk.Frame(self.root, relief=tk.RIDGE, borderwidth=2)
        password_frame.pack(pady=20, padx=50, fill=tk.X)
        
        tk.Label(password_frame, text="パスワード認証", font=("Arial", 14, "bold"), 
                bg="lightblue").pack(fill=tk.X, pady=5)
        
        input_frame = tk.Frame(password_frame)
        input_frame.pack(pady=15)
        
        tk.Label(input_frame, text="パスワード:", font=("Arial", 12)).pack(side=tk.LEFT, padx=5)
        password_entry = tk.Entry(input_frame, width=20, font=("Arial", 12), show='*')
        password_entry.pack(side=tk.LEFT, padx=5)
        
        def check_password():
            password = password_entry.get()
            if not password:
                messagebox.showerror("エラー", "パスワードを入力してください")
                return
            
            password_hash = hashlib.sha256(password.encode('utf-8')).hexdigest()
            if password_hash == PASSWORD_HASH:
                self.show_correction_menu()
            else:
                messagebox.showerror("エラー", "パスワードが正しくありません")
                password_entry.delete(0, tk.END)
        
        tk.Button(input_frame, text="認証", command=check_password,
                 font=("Arial", 12), bg="blue", fg="white").pack(side=tk.LEFT, padx=5)
        
        # マスターキーカード認証フレーム
        card_frame = tk.Frame(self.root, relief=tk.RIDGE, borderwidth=2)
        card_frame.pack(pady=20, padx=50, fill=tk.X)
        
        tk.Label(card_frame, text="マスターキーカード認証", font=("Arial", 14, "bold"), 
                bg="lightgreen").pack(fill=tk.X, pady=5)
        
        reader_frame = tk.Frame(card_frame)
        reader_frame.pack(pady=10)
        
        tk.Label(reader_frame, text="使用するリーダー:", font=("Arial", 11)).pack(side=tk.LEFT, padx=5)
        reader_var = tk.StringVar(value="class")
        tk.Radiobutton(reader_frame, text="授業用", variable=reader_var, 
                      value="class", font=("Arial", 10)).pack(side=tk.LEFT)
        tk.Radiobutton(reader_frame, text="会議用", variable=reader_var, 
                      value="meeting", font=("Arial", 10)).pack(side=tk.LEFT)
        
        status_label = tk.Label(card_frame, text="マスターキーカードをかざしてください...",
                              font=("Arial", 12), fg="blue")
        status_label.pack(pady=15)
        
        # カード監視用の変数
        auth_state = {'authenticated': False, 'monitoring': True}
        
        def check_master_card():
            if not auth_state['monitoring']:
                return
            
            selected_reader = self.card_reader_manager.class_reader if reader_var.get() == "class" else self.card_reader_manager.meeting_reader
            
            if selected_reader and self.card_reader_manager.is_card_present(selected_reader):
                connection = self.card_reader_manager.connect_to_card(selected_reader)
                if connection:
                    uid = self.card_reader_manager.get_card_uid(connection)
                    if uid:
                        if self.db_manager.is_master_key(uid):
                            auth_state['authenticated'] = True
                            auth_state['monitoring'] = False
                            status_label.config(text="認証成功！", fg="green")
                            self.sound_manager.play_beep("success")
                            self.root.after(500, self.show_correction_menu)
                        else:
                            status_label.config(text="このカードはマスターキーではありません", fg="red")
                            self.sound_manager.play_beep("error")
                            self.root.after(2000, lambda: status_label.config(
                                text="マスターキーカードをかざしてください...", fg="blue"))
                    self.card_reader_manager.disconnect(connection)
            
            if auth_state['monitoring']:
                self.root.after(500, check_master_card)
        
        check_master_card()
        
        # マスターキー管理ボタン
        tk.Button(card_frame, text="マスターキー管理", command=lambda: self.show_master_key_management(auth_state),
                 font=("Arial", 10), bg="gray", fg="white").pack(pady=10)
        
        tk.Button(self.root, text="戻る", command=lambda: self.cancel_auth(auth_state),
                 font=("Arial", 12)).pack(pady=20)
    
    def cancel_auth(self, auth_state):
        """認証をキャンセル"""
        auth_state['monitoring'] = False
        self.show_menu_callback()
    
    def show_correction_menu(self):
        """打刻修正メニュー画面"""
        for widget in self.root.winfo_children():
            widget.destroy()
        
        tk.Label(self.root, text="打刻修正メニュー", font=("Arial", 18, "bold")).pack(pady=30)
        
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=20)
        
        tk.Button(button_frame, text="打刻登録", 
                 command=self.show_correction_register,
                 font=("Arial", 14), bg="green", fg="white", 
                 width=20, height=3).pack(pady=15)
        
        tk.Button(button_frame, text="打刻削除", 
                 command=self.show_correction_delete,
                 font=("Arial", 14), bg="red", fg="white", 
                 width=20, height=3).pack(pady=15)
        
        tk.Button(self.root, text="戻る", command=self.show_menu_callback,
                 font=("Arial", 12)).pack(pady=20)
    
    def show_correction_register(self):
        """打刻登録画面"""
        for widget in self.root.winfo_children():
            widget.destroy()
        
        tk.Label(self.root, text="打刻登録", font=("Arial", 18, "bold")).pack(pady=20)
        
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
        
        tk.Button(self.root, text="戻る", command=self.show_correction_menu,
                 font=("Arial", 12)).pack(pady=10)
    
    def show_correction_delete(self):
        """打刻削除画面"""
        for widget in self.root.winfo_children():
            widget.destroy()
        
        tk.Label(self.root, text="打刻削除", font=("Arial", 18, "bold")).pack(pady=10)
        
        # フィルター入力フレーム
        filter_frame = tk.Frame(self.root)
        filter_frame.pack(pady=10)
        
        tk.Label(filter_frame, text="種別:", font=("Arial", 11)).grid(row=0, column=0, padx=5, pady=5)
        table_var = tk.StringVar(value="time_records")
        tk.Radiobutton(filter_frame, text="授業用", variable=table_var, 
                      value="time_records", font=("Arial", 10)).grid(row=0, column=1, sticky='w')
        tk.Radiobutton(filter_frame, text="会議用", variable=table_var, 
                      value="meeting_records", font=("Arial", 10)).grid(row=0, column=2, sticky='w')
        
        tk.Label(filter_frame, text="日付:", font=("Arial", 11)).grid(row=1, column=0, padx=5, pady=5)
        date_entry = tk.Entry(filter_frame, width=15, font=("Arial", 11))
        date_entry.grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky='w')
        
        today = datetime.now(JST).strftime("%Y-%m-%d")
        date_entry.insert(0, today)
        
        # 打刻一覧テーブル
        table_frame = tk.Frame(self.root)
        table_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(table_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        columns = ('ID', '講師名', '種別', '時刻')
        tree = ttk.Treeview(table_frame, columns=columns, show='headings',
                           yscrollcommand=scrollbar.set)
        
        tree.heading('ID', text='ID')
        tree.heading('講師名', text='講師名')
        tree.heading('種別', text='種別')
        tree.heading('時刻', text='時刻')
        
        tree.column('ID', width=60)
        tree.column('講師名', width=150)
        tree.column('種別', width=80)
        tree.column('時刻', width=180)
        
        tree.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=tree.yview)
        
        count_label = tk.Label(self.root, text="", font=("Arial", 10))
        count_label.pack(pady=5)
        
        def load_records():
            """打刻記録を読み込み"""
            for item in tree.get_children():
                tree.delete(item)
            
            date_str = date_entry.get().strip()
            table_name = table_var.get()
            
            try:
                datetime.strptime(date_str, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("エラー", "日付形式が正しくありません (YYYY-MM-DD)")
                return
            
            records = self.db_manager.get_date_records_with_id(date_str, table_name)
            
            for record_id, name, record_type, timestamp in records:
                action = "出勤" if record_type == "IN" else "退勤"
                tree.insert('', tk.END, values=(record_id, name, action, timestamp))
            
            count_label.config(text=f"記録数: {len(records)}件")
        
        def delete_selected():
            """選択された打刻記録を削除"""
            selected = tree.selection()
            if not selected:
                messagebox.showerror("エラー", "削除する記録を選択してください")
                return
            
            # 複数選択対応
            items = []
            for item in selected:
                values = tree.item(item)['values']
                items.append({
                    'id': values[0],
                    'name': values[1],
                    'type': values[2],
                    'time': values[3]
                })
            
            # 確認ダイアログ
            if len(items) == 1:
                msg = f"以下の記録を削除しますか？\n\n講師名: {items[0]['name']}\n種別: {items[0]['type']}\n時刻: {items[0]['time']}"
            else:
                msg = f"{len(items)}件の記録を削除しますか？"
            
            if not messagebox.askyesno("確認", msg):
                return
            
            # 削除実行
            table_name = table_var.get()
            success_count = 0
            
            for item in items:
                if self.db_manager.delete_attendance_record(item['id'], table_name):
                    success_count += 1
            
            if success_count == len(items):
                messagebox.showinfo("成功", f"{success_count}件の記録を削除しました")
                load_records()
            else:
                messagebox.showwarning("警告", f"{success_count}/{len(items)}件の記録を削除しました")
                load_records()
        
        # ボタンフレーム
        btn_frame = tk.Frame(filter_frame)
        btn_frame.grid(row=2, column=0, columnspan=3, pady=10)
        
        tk.Button(btn_frame, text="表示", command=load_records,
                 font=("Arial", 11), bg="blue", fg="white", width=10).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="削除", command=delete_selected,
                 font=("Arial", 11), bg="red", fg="white", width=10).pack(side=tk.LEFT, padx=5)
        
        # 初期表示
        load_records()
        
        tk.Button(self.root, text="戻る", command=self.show_correction_menu,
                 font=("Arial", 12)).pack(pady=10)
    
    def show_master_key_management(self, auth_state):
        """マスターキー管理画面（パスワード認証必要）"""
        auth_state['monitoring'] = False
        
        password = simpledialog.askstring("パスワード入力", 
            "マスターキー管理にはパスワードが必要です:", show='*')
        
        if password is None:
            self.show_correction_auth()
            return
        
        password_hash = hashlib.sha256(password.encode('utf-8')).hexdigest()
        if password_hash != PASSWORD_HASH:
            messagebox.showerror("エラー", "パスワードが正しくありません")
            self.show_correction_auth()
            return
        
        for widget in self.root.winfo_children():
            widget.destroy()
        
        tk.Label(self.root, text="マスターキー管理", font=("Arial", 18, "bold")).pack(pady=10)
        
        # マスターキー一覧
        table_frame = tk.Frame(self.root)
        table_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(table_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        columns = ('カードUID', '説明', '登録日時', '状態')
        tree = ttk.Treeview(table_frame, columns=columns, show='headings',
                           yscrollcommand=scrollbar.set)
        
        tree.heading('カードUID', text='カードUID')
        tree.heading('説明', text='説明')
        tree.heading('登録日時', text='登録日時')
        tree.heading('状態', text='状態')
        
        tree.column('カードUID', width=180)
        tree.column('説明', width=150)
        tree.column('登録日時', width=150)
        tree.column('状態', width=80)
        
        def refresh_list():
            for item in tree.get_children():
                tree.delete(item)
            
            master_keys = self.db_manager.get_master_keys()
            for key in master_keys:
                status = "有効" if key['is_active'] == 1 else "無効"
                tree.insert('', tk.END, values=(
                    key['card_uid'],
                    key['description'] or '',
                    key['created_at'],
                    status
                ))
        
        refresh_list()
        
        tree.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=tree.yview)
        
        # ボタンフレーム
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)
        
        def add_master_key():
            self.show_add_master_key(refresh_list)
        
        def delete_master_key():
            selected = tree.selection()
            if not selected:
                messagebox.showerror("エラー", "削除するマスターキーを選択してください")
                return
            
            item = tree.item(selected[0])
            card_uid = item['values'][0]
            
            if messagebox.askyesno("確認", f"マスターキー\n{card_uid}\nを削除しますか？"):
                if self.db_manager.delete_master_key(card_uid):
                    messagebox.showinfo("成功", "マスターキーを削除しました")
                    refresh_list()
                else:
                    messagebox.showerror("エラー", "削除に失敗しました")
        
        tk.Button(btn_frame, text="追加", command=add_master_key,
                 font=("Arial", 12), bg="green", fg="white", width=10).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="削除", command=delete_master_key,
                 font=("Arial", 12), bg="red", fg="white", width=10).pack(side=tk.LEFT, padx=5)
        
        tk.Button(self.root, text="戻る", command=self.show_correction_auth,
                 font=("Arial", 12)).pack(pady=10)
    
    def show_add_master_key(self, refresh_callback):
        """マスターキー追加画面"""
        add_window = tk.Toplevel(self.root)
        add_window.title("マスターキー追加")
        add_window.geometry("450x300")
        
        tk.Label(add_window, text="マスターキー追加", font=("Arial", 16, "bold")).pack(pady=10)
        
        reader_frame = tk.Frame(add_window)
        reader_frame.pack(pady=5)
        
        tk.Label(reader_frame, text="使用するリーダー:", font=("Arial", 11)).pack(side=tk.LEFT, padx=5)
        reader_var = tk.StringVar(value="class")
        tk.Radiobutton(reader_frame, text="授業用", variable=reader_var, 
                      value="class", font=("Arial", 10)).pack(side=tk.LEFT)
        tk.Radiobutton(reader_frame, text="会議用", variable=reader_var, 
                      value="meeting", font=("Arial", 10)).pack(side=tk.LEFT)
        
        status_label = tk.Label(add_window, text="カードをかざしてください...",
                              font=("Arial", 12), fg="blue")
        status_label.pack(pady=10)
        
        input_frame = tk.Frame(add_window)
        input_frame.pack(pady=10)
        
        tk.Label(input_frame, text="カードUID:").grid(row=0, column=0, padx=5, pady=5)
        uid_entry = tk.Entry(input_frame, width=30, state='readonly')
        uid_entry.grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(input_frame, text="説明:").grid(row=1, column=0, padx=5, pady=5)
        desc_entry = tk.Entry(input_frame, width=30)
        desc_entry.grid(row=1, column=1, padx=5, pady=5)
        
        detected_uid = {'uid': None}
        
        def check_card():
            if add_window.winfo_exists():
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
                            status_label.config(text="カード検出！説明を入力してください", fg="green")
                            self.sound_manager.play_beep("card_detected")
                        self.card_reader_manager.disconnect(connection)
                add_window.after(500, check_card)
        
        check_card()
        
        def register():
            uid = uid_entry.get().strip()
            description = desc_entry.get().strip()
            
            if not uid:
                messagebox.showerror("エラー", "カードをかざしてください")
                return
            
            if self.db_manager.add_master_key(uid, description):
                messagebox.showinfo("成功", "マスターキーを登録しました")
                add_window.destroy()
                refresh_callback()
            else:
                messagebox.showerror("エラー", "登録に失敗しました（既に登録済みの可能性があります）")
        
        btn_frame = tk.Frame(add_window)
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="登録", command=register, 
                 font=("Arial", 12), bg="green", fg="white", width=10).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="キャンセル", command=add_window.destroy,
                 font=("Arial", 12), width=10).pack(side=tk.LEFT, padx=5)
