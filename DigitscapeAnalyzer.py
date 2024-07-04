import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
import os


class DigitscapeAnalyzer:
    def __init__(self):
        self.file_path = None
        self.data = None
        self.stats = None
        self.is_normal = None
        self.random_numbers = None

    def load_data(self):
        if not self.file_path or not os.path.exists(self.file_path):
            print("Файл не выбран или не существует.")
            return

        try:
            self.data = pd.read_excel(self.file_path)
            print(f"Загружено {len(self.data)} строк из файла.")

            print("Первые 5 строк загруженных данных:")
            print(self.data.head())

            print("Типы данных:")
            print(self.data.dtypes)

            # Преобразование данных
            for column in self.data.columns:
                self.data[column] = self.data[column].astype(str)
                self.data[column] = self.data[column].str.replace(' ', '').str.replace(',', '.')
                self.data[column] = self.data[column].str.replace(r'[^\d.-]', '', regex=True)
                self.data[column] = pd.to_numeric(self.data[column], errors='coerce')

            print("Первые 5 строк после преобразования:")
            print(self.data.head())

            print(f"Количество строк с NaN после преобразования: {self.data.isna().sum().sum()}")

            # Удаление строк с NaN
            self.data.dropna(inplace=True)
            print(f"После удаления нечисловых значений осталось {len(self.data)} записей.")

            print("Первые 5 строк после удаления NaN:")
            print(self.data.head())

        except Exception as e:
            print(f"Ошибка при загрузке данных: {e}")
            raise

    def preprocess_data(self):
        if self.data is None or len(self.data) == 0:
            print("Нет данных для обработки.")
            return

        Q1 = self.data['numbers'].quantile(0.25)
        Q3 = self.data['numbers'].quantile(0.75)
        IQR = Q3 - Q1
        self.data = self.data[~((self.data['numbers'] < (Q1 - 1.5 * IQR)) | (self.data['numbers'] > (Q3 + 1.5 * IQR)))]
        print(f"После удаления выбросов осталось {len(self.data)} записей.")

    def calculate_statistics(self):
        if self.data is None or len(self.data) == 0:
            print("Нет данных для расчета статистики.")
            return None

        self.stats = self.data['numbers'].describe()
        self.stats['mode'] = self.data['numbers'].mode().values[0]
        self.stats['skewness'] = self.data['numbers'].skew()
        self.stats['kurtosis'] = self.data['numbers'].kurtosis()
        print(self.stats)
        return self.stats

    def plot_histogram(self):
        if self.data is None or len(self.data) == 0:
            print("Нет данных для построения гистограммы.")
            return

        plt.figure(figsize=(10, 6))
        sns.histplot(data=self.data, x='numbers', kde=True)
        plt.title('Гистограмма распределения данных')
        plt.xlabel('Значения')
        plt.ylabel('Частота')
        plt.savefig('histogram.png')
        plt.close()

    def check_normality(self):
        if self.data is None or len(self.data) == 0:
            print("Нет данных для проверки нормальности.")
            return None

        _, p_value = stats.normaltest(self.data['numbers'])
        self.is_normal = p_value >= 0.05
        print(f"p-значение теста на нормальность: {p_value}")
        print(f"Распределение {'является' if self.is_normal else 'не является'} нормальным.")
        return self.is_normal

    def generate_random_numbers(self, count=37):
        if self.data is None or len(self.data) == 0:
            print("Нет данных для генерации случайных чисел.")
            return None

        self.random_numbers = np.random.choice(self.data['numbers'], size=count, replace=False)
        return self.random_numbers

    def check_random_numbers_uniformity(self):
        if self.random_numbers is None:
            print("Сначала сгенерируйте случайные числа.")
            return None

        _, p_value = stats.kstest(self.random_numbers, 'uniform',
                                  args=(self.random_numbers.min(), self.random_numbers.max()))
        is_uniform = p_value >= 0.05
        print(f"p-значение теста на равномерность: {p_value}")
        print(f"Распределение сгенерированных чисел {'является' if is_uniform else 'не является'} равномерным.")
        return is_uniform

    def plot_boxplot(self):
        if self.data is None or len(self.data) == 0:
            print("Нет данных для построения box plot.")
            return

        plt.figure(figsize=(10, 6))
        sns.boxplot(x=self.data['numbers'])
        plt.title('Box plot данных')
        plt.xlabel('Значения')
        plt.savefig('boxplot.png')
        plt.close()

    def plot_scatter(self):
        if self.data is None or len(self.data) == 0:
            print("Нет данных для построения диаграммы рассеяния.")
            return

        plt.figure(figsize=(10, 6))
        sns.scatterplot(x=range(len(self.data)), y='numbers', data=self.data)
        plt.title('Диаграмма рассеяния данных')
        plt.xlabel('Индекс')
        plt.ylabel('Значения')
        plt.savefig('scatter.png')
        plt.close()

    def generate_report(self):
        if self.stats is None:
            print("Сначала выполните статистический анализ.")
            return

        report = "Отчет по анализу данных:\n\n"
        report += f"1. Статистические характеристики:\n{self.stats}\n\n"
        report += f"2. Нормальность распределения: {'Да' if self.is_normal else 'Нет'}\n\n"
        report += "3. Визуализация:\n"
        report += "   - Гистограмма: histogram.png\n"
        report += "   - Box plot: boxplot.png\n"
        report += "   - Диаграмма рассеяния: scatter.png\n\n"
        report += f"4. Сгенерированные случайные числа: {self.random_numbers}\n"
        report += f"   Равномерность распределения: {self.check_random_numbers_uniformity()}\n\n"
        report += "5. Выводы и рекомендации:\n"
        report += "   - " + self.generate_conclusions()

        with open('report.txt', 'w') as f:
            f.write(report)

        print("Отчет сохранен в файл 'report.txt'")

    def generate_conclusions(self):
        conclusions = []
        if not self.is_normal:
            conclusions.append("Распределение данных отклоняется от нормального. "
                               "Рекомендуется использовать непараметрические методы для дальнейшего анализа.")
        if self.stats['skewness'] > 1 or self.stats['skewness'] < -1:
            conclusions.append("Наблюдается значительная асимметрия распределения. "
                               "Рекомендуется исследовать причины и, возможно, применить преобразование данных.")
        if self.stats['kurtosis'] > 3:
            conclusions.append("Распределение имеет тяжелые хвосты. "
                               "Следует обратить внимание на выбросы и их влияние на анализ.")

        return " ".join(
            conclusions) if conclusions else "Значительных отклонений от математической модели не обнаружено."

    def run_analysis(self):
        try:
            self.load_data()
            self.preprocess_data()
            self.calculate_statistics()
            self.plot_histogram()
            self.check_normality()
            self.generate_random_numbers()
            self.check_random_numbers_uniformity()
            self.plot_boxplot()
            self.plot_scatter()
            self.generate_report()
            print("Анализ успешно завершен.")
        except Exception as e:
            print(f"Ошибка при выполнении анализа: {e}")