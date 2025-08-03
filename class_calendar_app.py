import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import json
import os
import shutil

class ClassCalendarApp:
    DAYS = ['月', '火', '水', '木', '金']
    PERIODS = [1, 2, 3, 4, 5, 6]
    DATA_FILE = 'timetable_data.json'
    DEFAULT_SEMESTER = "Default"

    def __init__(self, root):
        self.root = root
        self.root.title("授業カレンダー")
        self.all_data = self.load_all_data()
        self.current_semester = self.all_data.get('current_semester', self.DEFAULT_SEMESTER)
        self.data = self.all_data['semesters'].get(self.current_semester, {})
        self.word_template_path = self.all_data.get('word_template_path', '')
        self.student_id = self.all_data.get('student_id', '')
        
        self.create_widgets()
        self.update_title()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def load_all_data(self):
        if not os.path.exists(self.DATA_FILE):
            return {"current_semester": self.DEFAULT_SEMESTER, "semesters": {self.DEFAULT_SEMESTER: {}}, "student_id": ""}
        
        with open(self.DATA_FILE, 'r', encoding='utf-8') as f:
            try:
                data = json.load(f)
                if "semesters" not in data:
                    return {"current_semester": self.DEFAULT_SEMESTER, "semesters": {self.DEFAULT_SEMESTER: data}, "student_id": ""}
                return data
            except json.JSONDecodeError:
                return {"current_semester": self.DEFAULT_SEMESTER, "semesters": {self.DEFAULT_SEMESTER: {}}, "student_id": ""}

    def save_all_data(self):
        self.all_data['semesters'][self.current_semester] = self.data
        self.all_data['word_template_path'] = self.word_template_path
        self.all_data['student_id'] = self.student_id
        with open(self.DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.all_data, f, ensure_ascii=False, indent=4)

    def on_closing(self):
        self.save_all_data()
        self.root.destroy()

    def update_title(self):
        self.root.title(f"授業カレンダー - {self.current_semester}")

    def create_widgets(self):
        # Menu
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="ファイル", menu=file_menu)
        file_menu.add_command(label="学期の管理...", command=self.open_semester_manager)
        file_menu.add_command(label="Wordテンプレートをアップロード...", command=self.upload_word_template)
        file_menu.add_command(label="学籍番号を設定...", command=self.set_student_id)
        file_menu.add_separator()
        file_menu.add_command(label="終了", command=self.on_closing)

        # Main Frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Header
        for c, day in enumerate(self.DAYS):
            ttk.Label(main_frame, text=day, font=('Helvetica', 12, 'bold')).grid(row=0, column=c + 1, padx=5, pady=5)
        for r, period in enumerate(self.PERIODS):
            ttk.Label(main_frame, text=f"{period}限", font=('Helvetica', 12, 'bold')).grid(row=r + 1, column=0, padx=5, pady=5)

        # Timetable cells
        self.cell_buttons = {}
        for r, period in enumerate(self.PERIODS):
            for c, day in enumerate(self.DAYS):
                key = f"{day}-{period}"
                btn = tk.Button(main_frame, text=self.get_cell_display_text(key), width=15, height=5,
                                command=lambda k=key: self.open_edit_window(k))
                btn.grid(row=r + 1, column=c + 1, padx=2, pady=2)
                self.cell_buttons[key] = btn
    
    def refresh_timetable(self):
        self.data = self.all_data['semesters'].get(self.current_semester, {})
        for key, btn in self.cell_buttons.items():
            btn.config(text=self.get_cell_display_text(key))
        self.update_title()

    def get_cell_display_text(self, key):
        class_info = self.data.get(key)
        if not class_info:
            return ""
        return f"{class_info.get('subject', '')}\n{class_info.get('teacher', '')}\n{class_info.get('classroom', '')}"

    def open_edit_window(self, key):
        edit_window = tk.Toplevel(self.root)
        edit_window.title("授業の編集")

        class_info = self.data.get(key, {})
        
        ttk.Label(edit_window, text="授業名:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        subject_entry = ttk.Entry(edit_window, width=30)
        subject_entry.grid(row=0, column=1, padx=10, pady=5)
        subject_entry.insert(0, class_info.get('subject', ''))

        ttk.Label(edit_window, text="教師:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        teacher_entry = ttk.Entry(edit_window, width=30)
        teacher_entry.grid(row=1, column=1, padx=10, pady=5)
        teacher_entry.insert(0, class_info.get('teacher', ''))

        ttk.Label(edit_window, text="教室:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        classroom_entry = ttk.Entry(edit_window, width=30)
        classroom_entry.grid(row=2, column=1, padx=10, pady=5)
        classroom_entry.insert(0, class_info.get('classroom', ''))

        note_path_var = tk.StringVar(value=class_info.get('note_path', ''))
        ttk.Label(edit_window, text="ノートファイル:").grid(row=3, column=0, padx=10, pady=5, sticky="w")
        note_label = ttk.Label(edit_window, textvariable=note_path_var, wraplength=200)
        note_label.grid(row=3, column=1, padx=10, pady=5, sticky="w")

        open_note_button = ttk.Button(edit_window, text="ノートを開く", command=lambda: self.open_note_file(note_path_var.get()))
        open_note_button.grid(row=4, column=0, padx=10, pady=10, sticky="w")

        def update_open_button_state(*args):
            if note_path_var.get() and os.path.exists(note_path_var.get()):
                open_note_button.config(state=tk.NORMAL)
            else:
                open_note_button.config(state=tk.DISABLED)

        note_path_var.trace_add("write", update_open_button_state)
        update_open_button_state()

        def select_note_file():
            note_path_var.trace_add("write", update_open_button_state)
        update_open_button_state()

        def select_note_file():
            filepath = filedialog.askopenfilename(
                title="ノートファイル (.md) を選択",
                filetypes=[("Markdown files", "*.md"), ("All files", "*.* אמיתי")])
            if filepath:
                note_path_var.set(filepath)

        ttk.Button(edit_window, text="ノートを選択...", command=select_note_file).grid(row=4, column=1, padx=10, pady=10, sticky="e")

        def save_and_close():
            new_info = {
                'subject': subject_entry.get(),
                'teacher': teacher_entry.get(),
                'classroom': classroom_entry.get(),
                'note_path': note_path_var.get()
            }
            if any(new_info.values()):
                self.data[key] = new_info
            elif key in self.data:
                del self.data[key]
            
            self.cell_buttons[key].config(text=self.get_cell_display_text(key))
            self.save_all_data()
            edit_window.destroy()

        button_frame = ttk.Frame(edit_window)
        button_frame.grid(row=5, column=0, columnspan=2, pady=10)

        ttk.Button(button_frame, text="保存", command=save_and_close).pack(side="left", padx=5)
        ttk.Button(button_frame, text="キャンセル", command=edit_window.destroy).pack(side="left", padx=5)
        
        def delete_class_and_close():
            if messagebox.askyesno("確認", "本当にこの授業を削除しますか？"):
                if key in self.data:
                    del self.data[key]
                    self.cell_buttons[key].config(text=self.get_cell_display_text(key))
                    self.save_all_data()
                edit_window.destroy()

        ttk.Button(button_frame, text="削除", command=delete_class_and_close).pack(side="left", padx=5)
        
        # --- 課題作成ボタン ---
        can_create_assignment = self.word_template_path and os.path.exists(self.word_template_path) and self.student_id
        assignment_button_state = tk.NORMAL if can_create_assignment else tk.DISABLED
        
        create_assignment_button = ttk.Button(button_frame, text="課題を作成", 
                                               command=lambda: self.create_assignment_from_template(key, edit_window),
                                               state=assignment_button_state)
        create_assignment_button.pack(side="right", padx=20)

        if not can_create_assignment:
            tooltip_text = ""
            if not self.word_template_path or not os.path.exists(self.word_template_path):
                tooltip_text += "・Wordテンプレートが未設定です\n"
            if not self.student_id:
                tooltip_text += "・学籍番号が未設定です"
            
            # 簡単なツールチップの実装
            def show_tooltip(event):
                tooltip = tk.Toplevel(edit_window)
                tooltip.wm_overrideredirect(True)
                tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
                label = tk.Label(tooltip, text=tooltip_text.strip(), justify='left',
                                 background="#ffffe0", relief='solid', borderwidth=1,
                                 font=("tahoma", "8", "normal"))
                label.pack(ipadx=1)
                create_assignment_button.bind("<Leave>", lambda e: tooltip.destroy())
            
            create_assignment_button.bind("<Enter>", show_tooltip)

    def open_note_file(self, filepath):
        if filepath and os.path.exists(filepath):
            os.startfile(filepath)
        else:
            messagebox.showerror("エラー", "ファイルが見つかりません。")

    def open_semester_manager(self):
        manager_window = tk.Toplevel(self.root)
        manager_window.title("学期の管理")
        manager_window.geometry("350x300")

        frame = ttk.Frame(manager_window, padding="10")
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="学期一覧:").pack(anchor="w")
        
        listbox = tk.Listbox(frame)
        listbox.pack(fill="both", expand=True, pady=5)
        
        semesters = list(self.all_data["semesters"].keys())
        for semester in semesters:
            listbox.insert(tk.END, semester)
        
        try:
            current_idx = semesters.index(self.current_semester)
            listbox.selection_set(current_idx)
            listbox.see(current_idx)
        except ValueError:
            pass

        def add_semester():
            new_name = simpledialog.askstring("新しい学期", "新しい学期の名前を入力してください:", parent=manager_window)
            if new_name and new_name not in self.all_data["semesters"]:
                self.all_data["semesters"][new_name] = {}
                listbox.insert(tk.END, new_name)
                self.save_all_data()
            elif new_name:
                messagebox.showwarning("警告", "その学期名はすでに存在します。", parent=manager_window)

        def switch_semester():
            selected_indices = listbox.curselection()
            if not selected_indices:
                return
            selected_semester = listbox.get(selected_indices[0])
            if selected_semester != self.current_semester:
                self.save_all_data() # Save current semester before switching
                self.current_semester = selected_semester
                self.all_data['current_semester'] = self.current_semester
                self.refresh_timetable()
                self.save_all_data() # Save the fact that we switched
            manager_window.destroy()

        def reset_semester():
            selected_indices = listbox.curselection()
            if not selected_indices:
                return
            selected_semester = listbox.get(selected_indices[0])
            if messagebox.askyesno("確認", f"'{selected_semester}' の時間割を本当にリセットしますか？\nこの操作は元に戻せません。", parent=manager_window):
                self.all_data["semesters"][selected_semester] = {}
                if selected_semester == self.current_semester:
                    self.refresh_timetable()
                self.save_all_data()

        def delete_semester():
            selected_indices = listbox.curselection()
            if not selected_indices:
                messagebox.showwarning("警告", "削除する学期を選択してください。", parent=manager_window)
                return

            if len(self.all_data["semesters"]) <= 1:
                messagebox.showerror("エラー", "最後の学期は削除できません。", parent=manager_window)
                return

            selected_semester = listbox.get(selected_indices[0])
            
            if messagebox.askyesno("確認", f"本当に学期 '{selected_semester}' を削除しますか？\nこの操作は元に戻せません。", parent=manager_window):
                del self.all_data["semesters"][selected_semester]
                listbox.delete(selected_indices[0])

                if self.current_semester == selected_semester:
                    self.current_semester = list(self.all_data["semesters"].keys())[0]
                    self.all_data['current_semester'] = self.current_semester
                    self.refresh_timetable()
                
                self.save_all_data()
                messagebox.showinfo("成功", f"学期 '{selected_semester}' を削除しました。", parent=manager_window)

        button_frame = ttk.Frame(frame)
        button_frame.pack(fill="x", pady=5)
        
        ttk.Button(button_frame, text="追加", command=add_semester).pack(side="left", expand=True, fill="x", padx=2)
        ttk.Button(button_frame, text="削除", command=delete_semester).pack(side="left", expand=True, fill="x", padx=2)
        ttk.Button(button_frame, text="リセット", command=reset_semester).pack(side="left", expand=True, fill="x", padx=2)
        
        ttk.Button(frame, text="この学期に切り替え", command=switch_semester).pack(fill="x", pady=5)
        ttk.Button(frame, text="閉じる", command=manager_window.destroy).pack(fill="x")

    def upload_word_template(self):
        filepath = filedialog.askopenfilename(
            title="Wordテンプレートファイル (.docx) を選択",
            filetypes=[("Word Document", "*.docx"), ("All files", "*.*")])
        if filepath:
            self.word_template_path = filepath
            self.save_all_data()
            messagebox.showinfo("成功", "テンプレートが保存されました。")

    def set_student_id(self):
        new_id = simpledialog.askstring("学籍番号の設定", "学籍番号を入力してください:",
                                        initialvalue=self.student_id, parent=self.root)
        if new_id:
            self.student_id = new_id
            self.save_all_data()
            messagebox.showinfo("成功", "学籍番号が保存されました。")

    def create_assignment_from_template(self, key, parent_window):
        class_info = self.data.get(key, {})
        subject = class_info.get('subject')

        if not subject:
            messagebox.showerror("エラー", "このセルに授業が設定されていません。", parent=parent_window)
            return

        assignment_num = simpledialog.askstring("課題情報", "第何回の課題ですか？（例: 1, 2, 最終）", parent=parent_window)
        if not assignment_num:
            return

        default_filename = f"{subject}_第{assignment_num}回_{self.student_id}.docx"
        
        save_path = filedialog.asksaveasfilename(
            title="課題ファイルを保存",
            initialfile=default_filename,
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")],
            parent=parent_window
        )

        if save_path:
            try:
                shutil.copy(self.word_template_path, save_path)
                messagebox.showinfo("成功", f"課題ファイルが作成されました:\n{save_path}", parent=parent_window)
                if messagebox.askyesno("確認", "作成したファイルを開きますか？", parent=parent_window):
                    os.startfile(save_path)
            except Exception as e:
                messagebox.showerror("エラー", f"ファイルの作成に失敗しました: {e}", parent=parent_window)


if __name__ == "__main__":
    root = tk.Tk()
    app = ClassCalendarApp(root)
    root.mainloop()
