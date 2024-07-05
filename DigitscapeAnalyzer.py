import pandas as pd
import numpy as np
from scipy import stats
import matplotlib.pyplot as plt
import seaborn as sns


class DigitscapeAnalyzer:
    def __init__(self):
        self.file_path = None
        self.data = None
        self.analysis_results = {}

    def load_data(self, file_path, selected_columns):
        self.file_path = file_path
        self.data = pd.read_excel(file_path, usecols=selected_columns)

    def analyze_distribution(self):
        for column in self.data.columns:
            # Проверка на нормальность распределения
            _, p_value_normal = stats.normaltest(self.data[column])

            # Проверка на равномерность распределения
            _, p_value_uniform = stats.kstest(self.data[column], 'uniform',
                                              args=(self.data[column].min(), self.data[column].max()))

            # Проверка на случайность последовательности (Run test)
            median = self.data[column].median()
            binary = (self.data[column] > median).astype(int)
            runs, p_value_run = stats.runs_test(binary)

            self.analysis_results[column] = {
                'normal_p_value': p_value_normal,
                'uniform_p_value': p_value_uniform,
                'run_test_p_value': p_value_run
            }

    def plot_distribution(self, column):
        plt.figure(figsize=(12, 6))
        sns.histplot(self.data[column], kde=True)
        plt.title(f'Распределение значений в столбце {column}')
        plt.savefig(f'distribution_{column}.png')
        plt.close()

    def generate_conclusions(self):
        conclusions = []
        for column, results in self.analysis_results.items():
            if results['normal_p_value'] < 0.05:
                conclusions.append(
                    f"Распределение в столбце {column} отклоняется от нормального (p={results['normal_p_value']:.4f}).")
            if results['uniform_p_value'] < 0.05:
                conclusions.append(
                    f"Распределение в столбце {column} не является равномерным (p={results['uniform_p_value']:.4f}).")
            if results['run_test_p_value'] < 0.05:
                conclusions.append(
                    f"Последовательность значений в столбце {column} может не быть случайной (p={results['run_test_p_value']:.4f}).")

        return conclusions if conclusions else ["Значительных отклонений от математической модели не обнаружено."]

    def run_analysis(self, selected_columns):
        self.load_data(self.file_path, selected_columns)
        self.analyze_distribution()
        for column in self.data.columns:
            self.plot_distribution(column)
        return self.generate_conclusions()