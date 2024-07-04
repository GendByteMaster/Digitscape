import os
import pandas as pd
import threading
from tkinter import filedialog, messagebox
from tkinter import *
from tkinter import ttk

class DigitscapeGUI:
    def __init__(self, master, analyzer):
        self.master = master
        self.analyzer = analyzer
        self.create_widgets()
        self.bind_keys()

    def create_widgets(self):
        main_frame = ttk.Frame(self.master, padding="20 20 20 20")
        main_frame.pack(fill=BOTH, expand=YES)

        title_label = ttk.Label(main_frame, text="Digitscape Analyzer", font=("Helvetica", 18, "bold"))
        title_label.pack(pady=(0, 20))

        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=X, pady=10)

        self.file_label = ttk.Label(file_frame, text="Файл не выбран", font=("Helvetica", 10))
        self.file_label.pack(side=LEFT, expand=YES)

        select_file_button = ttk.Button(file_frame, text="Выбрать файл", command=self.select_file,
                                        style="Outline.TButton")
        select_file_button.pack(side=RIGHT)

        self.column_listbox = Listbox(main_frame, selectmode=MULTIPLE)
        self.column_listbox.pack(fill=X, pady=10)

        self.progress_bar = ttk.Progressbar(main_frame, mode="indeterminate",
                                            style="success.Striped.Horizontal.TProgressbar")
        self.progress_bar.pack(fill=X, pady=20)

        self.run_analysis_button = ttk.Button(main_frame, text="Запустить анализ", command=self.run_analysis,
                                              style="success.TButton")
        self.run_analysis_button.pack(pady=10)

        self.view_report_button = ttk.Button(main_frame, text="Просмотреть отчет", command=self.view_report,
                                             style="info.TButton", state="disabled")
        self.view_report_button.pack(pady=10)

        self.status_label = ttk.Label(main_frame, text="Готов к работе", font=("Helvetica", 10))
        self.status_label.pack(pady=(20, 0))

    def bind_keys(self):
        self.master.bind('<Control-o>', self.select_file)
        self.master.bind('<Control-r>', self.run_analysis)
        self.master.bind('<Control-v>', self.view_report)

        self.master.bind('<Escape>', self.exit_program)

    def exit_program(self, event=None):
        if messagebox.askokcancel("Выход", "Вы уверены, что хотите выйти из программы?"):
            self.master.quit()
    def select_file(self, event=None):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.analyzer.file_path = file_path
            self.file_label.config(text=f"Выбран: {os.path.basename(file_path)}")
            self.status_label.config(text="Файл выбран. Готов к анализу.")
            self.load_columns()

    def load_columns(self):
        try:
            df = pd.read_excel(self.analyzer.file_path)
            self.column_listbox.delete(0, END)
            for column in df.columns:
                self.column_listbox.insert(END, column)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить столбцы: {e}")

    def run_analysis(self, event=None):
        if not self.analyzer.file_path:
            messagebox.showerror("Ошибка", "Пожалуйста, выберите файл перед запуском анализа.")
            return

        selected_columns = [self.column_listbox.get(i) for i in self.column_listbox.curselection()]
        if not selected_columns:
            messagebox.showerror("Ошибка", "Пожалуйста, выберите хотя бы один столбец для анализа.")
            return

        self.run_analysis_button.config(state="disabled")
        self.progress_bar.start(10)
        self.status_label.config(text="Выполняется анализ...")

        thread = threading.Thread(target=self.run_analysis_thread, args=(selected_columns,))
        thread.start()

    def run_analysis_thread(self, selected_columns):
        try:
            self.analyzer.run_analysis(selected_columns)
            self.master.after(0, self.analysis_complete)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка при выполнении анализа: {e}")

    def analysis_complete(self):
        self.run_analysis_button.config(state="normal")
        self.progress_bar.stop()
        self.status_label.config(text="Анализ завершен")
        self.view_report_button.config(state="normal")

    def view_report(self, event=None):
        if not self.analyzer.file_path:
            messagebox.showerror("Ошибка", "Пожалуйста, выберите файл перед просмотром отчета.")
            return

        try:
            self.analyzer.view_report()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка при просмотре отчета: {e}")