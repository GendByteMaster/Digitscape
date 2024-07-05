import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd


class DigitscapeGUI:
    def __init__(self, root, analyzer):
        self.root = root
        self.root.geometry("700x500")
        self.analyzer = analyzer
        self.root.title("Digitscape Analyzer")
        self.create_widgets()

    def create_widgets(self):
        # Main frame
        main_frame = ctk.CTkFrame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # File selection
        file_frame = ctk.CTkFrame(main_frame)
        file_frame.pack(fill="x", pady=10)
        self.file_label = ctk.CTkLabel(file_frame, text="Выберите файл с данными:")
        self.file_label.pack(side="left")
        self.file_entry = ctk.CTkEntry(file_frame, width=400)
        self.file_entry.pack(side="left", padx=5)
        self.file_button = ctk.CTkButton(file_frame, text="Выбрать файл", command=self.select_file)
        self.file_button.pack(side="left")

        # Column selection
        column_frame = ctk.CTkFrame(main_frame)
        column_frame.pack(fill="x", pady=10)
        self.column_label = ctk.CTkLabel(column_frame, text="Выберите столбец для анализа:")
        self.column_label.pack(side="left")
        self.column_entry = ctk.CTkEntry(column_frame, width=400)
        self.column_entry.pack(side="left", padx=5)

        # Analyze button
        self.analyze_button = ctk.CTkButton(main_frame, text="Анализировать", command=self.run_analysis)
        self.analyze_button.pack(pady=10)

        # Results area
        self.results_text = ctk.CTkTextbox(main_frame, height=300, width=600)
        self.results_text.pack(pady=10)

    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, file_path)

            # Получаем имена столбцов из Excel-файла
            try:
                df = pd.read_excel(file_path, nrows=0)
                column_names = df.columns.tolist()
                self.column_entry.delete(0, "end")
                self.column_entry.insert(0, column_names[0])  # Вставляем первый столбец по умолчанию
            except Exception as e:
                messagebox.showwarning("Предупреждение", f"Не удалось прочитать имена столбцов из файла: {str(e)}")

    def run_analysis(self):
        try:
            file_path = self.file_entry.get()
            column_name = self.column_entry.get()

            if not file_path or not column_name:
                raise ValueError("Выберите файл и укажите столбец для анализа")

            self.analyzer.load_data(file_path, column_name)
            self.analyzer.run_analysis()
            report = self.analyzer.generate_report()

            self.results_text.delete("1.0", "end")
            self.results_text.insert("end", report)

            messagebox.showinfo("Анализ завершен", "Анализ успешно завершен. Результаты отображены в окне приложения.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка при выполнении анализа: {str(e)}")