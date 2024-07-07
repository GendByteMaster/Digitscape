import customtkinter as ctk
import seaborn as sns
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends._backend_tk import NavigationToolbar2Tk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import asyncio
import threading

sns.set(style="darkgrid")

class DigitscapeGUI:
    def __init__(self, root, analyzer):
        """
        Инициализирует графический интерфейс Digitscape.

        :param root: корневой объект tkinter
        :param analyzer: экземпляр класса DigitscapeAnalyzer
        """
        self.root = root
        self.root.geometry("700x500")
        self.analyzer = analyzer
        self.root.title("Digitscape Analyzer")
        self.create_widgets()
        self.bind_events()
        self.loop = asyncio.get_event_loop()

    def create_widgets(self):
        """Создает и размещает виджеты в главном окне."""
        # Основная рамка
        main_frame = ctk.CTkFrame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Выбор файла
        file_frame = ctk.CTkFrame(main_frame)
        file_frame.pack(fill="x", pady=10)
        self.file_label = ctk.CTkLabel(file_frame, text="Выберите файл с данными:")
        self.file_label.pack(side="left")
        self.file_entry = ctk.CTkEntry(file_frame, width=400)
        self.file_entry.pack(side="left", padx=5)
        self.file_button = ctk.CTkButton(file_frame, text="Выбрать файл", command=self.select_file)
        self.file_button.pack(side="left")

        # Выбор столбца
        column_frame = ctk.CTkFrame(main_frame)
        column_frame.pack(fill="x", pady=10)
        self.column_label = ctk.CTkLabel(column_frame, text="Выберите столбец для анализа:")
        self.column_label.pack(side="left")
        self.column_entry = ctk.CTkEntry(column_frame, width=400)
        self.column_entry.pack(side="left", padx=5)

        # Кнопка анализа
        self.analyze_button = ctk.CTkButton(main_frame, text="Анализировать", command=self.run_analysis)
        self.analyze_button.pack(pady=10)

        # Область результатов
        self.results_text = ctk.CTkTextbox(main_frame, height=300, width=600)
        self.results_text.pack(pady=10)

    def bind_events(self):
        """Привязывает события к виджетам."""
        self.root.bind("<Control-o>", lambda event: self.select_file())
        self.root.bind("<Control-a>", lambda event: self.run_analysis())
        self.root.bind("<Return>", lambda event: self.run_analysis())
        self.root.bind("<Escape>", lambda event: self.root.destroy())
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def on_close(self):
        """Обрабатывает событие закрытия окна."""
        if messagebox.askokcancel("Выход", "Вы уверены, что хотите выйти?"):
            self.root.destroy()

    def select_file(self):
        """Открывает диалоговое окно для выбора файла и обновляет интерфейс."""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, file_path)

            try:
                # Читаем первую строку файла как заголовки
                df = pd.read_excel(file_path, nrows=0)

                # Преобразуем все имена столбцов в строки
                column_names = [str(col) for col in df.columns]

                self.column_entry.delete(0, "end")
                self.column_entry.insert(0, ", ".join(column_names))

                messagebox.showinfo("Успех", f"Файл успешно загружен. Найдено {len(column_names)} столбцов.")
            except Exception as e:
                messagebox.showwarning("Предупреждение", f"Не удалось прочитать имена столбцов из файла: {str(e)}")

    def run_analysis(self):
        """Запускает процесс анализа данных."""
        file_path = self.file_entry.get()
        column_names = [col.strip() for col in self.column_entry.get().split(',')]

        if not file_path or not column_names:
            messagebox.showerror("Ошибка", "Выберите файл и укажите столбец(ы) для анализа")
            return

        # Запускаем асинхронный анализ в отдельном потоке
        threading.Thread(target=self.async_run_analysis, args=(file_path, column_names)).start()

    def async_run_analysis(self, file_path, column_names):
        """
        Выполняет асинхронный анализ данных.

        :param file_path: путь к файлу с данными
        :param column_names: список названий столбцов для анализа
        """
        async def analysis_task():
            try:
                # Загрузка данных
                await self.analyzer.load_data(file_path, column_names)

                # Выполнение анализа
                conclusions = await self.analyzer.run_analysis(column_names)

                # Обновление UI должно происходить в главном потоке
                self.root.after(0, self.update_results, conclusions)
                self.root.after(0, self.show_graphs, column_names)
                self.root.after(0, lambda: messagebox.showinfo("Анализ завершен",
                                                               "Анализ успешно завершен. Результаты отображены в окне приложения."))
            except Exception as e:
                error_message = f"Произошла ошибка при выполнении анализа: {str(e)}"
                self.root.after(0, lambda: messagebox.showerror("Ошибка", error_message))

        asyncio.run(analysis_task())

    def update_results(self, conclusions):
        """
        Обновляет текстовое поле с результатами анализа.

        :param conclusions: список выводов анализа
        """
        self.results_text.delete("1.0", "end")
        self.results_text.insert("end", "\n".join(conclusions))

    def show_graphs(self, column_names):
        """
        Отображает графики распределения для выбранных столбцов.

        :param column_names: список названий столбцов для визуализации
        """
        graph_window = ctk.CTkToplevel(self.root)
        graph_window.title("Графики распределения")
        graph_window.geometry("1920x1080")

        for i, column_name in enumerate(column_names):
            fig, ax = plt.subplots()
            sns.histplot(self.analyzer.data[column_name], bins=20, kde=True, ax=ax)
            ax.set_title(f"Гистограмма распределения для столбца '{column_name}'", fontsize=16)
            ax.set_xlabel(column_name, fontsize=14)
            ax.set_ylabel("Количество", fontsize=14)
            ax.grid(True)

            ax.tick_params(axis='x', labelsize=12)
            ax.tick_params(axis='y', labelsize=12)

            canvas = FigureCanvasTkAgg(fig, master=graph_window)
            canvas.draw()
            canvas.get_tk_widget().pack(side="top", fill="both", expand=True)

            toolbar = NavigationToolbar2Tk(canvas, graph_window)
            toolbar.update()
            canvas.get_tk_widget().pack(side="top", fill="both", expand=True)

        graph_window.mainloop()