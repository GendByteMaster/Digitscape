import pandas as pd
import numpy as np
from scipy import stats
from scipy.optimize import curve_fit
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.cluster import KMeans
import asyncio
from concurrent.futures import ThreadPoolExecutor

sns.set(style="darkgrid")


class DigitscapeAnalyzer:
    def __init__(self):
        self.file_path = None
        self.data = None
        self.analysis_results = {}
        self.executor = ThreadPoolExecutor()

    async def load_data(self, file_path, selected_columns):
        """
        Асинхронно загружает данные из Excel-файла.

        :param file_path: путь к файлу Excel
        :param selected_columns: список выбранных столбцов или строка с названиями столбцов через запятую
        """
        self.file_path = file_path
        try:
            all_data = await asyncio.get_event_loop().run_in_executor(
                self.executor, pd.read_excel, file_path
            )

            print(f"Доступные столбцы в файле: {all_data.columns.tolist()}")

            if isinstance(selected_columns, str):
                selected_columns = [col.strip() for col in selected_columns.split(',')]

            valid_columns = [col for col in selected_columns if col in all_data.columns]
            invalid_columns = set(selected_columns) - set(valid_columns)

            if not valid_columns:
                raise ValueError(f"Ни один из выбранных столбцов не найден в файле. "
                                 f"Недействительные столбцы: {', '.join(invalid_columns)}")

            self.data = all_data[valid_columns]

            non_numeric_columns = [col for col in valid_columns if not pd.api.types.is_numeric_dtype(self.data[col])]
            if non_numeric_columns:
                raise ValueError(f"Следующие столбцы содержат нечисловые данные: {', '.join(non_numeric_columns)}")

            print(f"Успешно загружены столбцы: {', '.join(valid_columns)}")
            if invalid_columns:
                print(f"Предупреждение: следующие столбцы не найдены и будут пропущены: {', '.join(invalid_columns)}")

        except Exception as e:
            raise ValueError(f"Ошибка при загрузке данных: {str(e)}")

    async def run_analysis(self, column_names):
        """
        Запускает анализ для выбранных столбцов.

        :param column_names: список названий столбцов для анализа
        :return: список выводов по результатам анализа
        """
        if self.data is None:
            raise ValueError("Данные не загружены. Пожалуйста, сначала загрузите данные с помощью метода load_data().")

        self.analysis_results = {}
        conclusions = []

        for column in column_names:
            if column not in self.data.columns:
                raise ValueError(f"Столбец '{column}' не найден в загруженных данных.")

            data = self.data[column].dropna()

            results = await asyncio.get_event_loop().run_in_executor(
                self.executor, self.analyze_column, data
            )
            self.analysis_results[column] = results

            conclusions.extend(await self.generate_column_conclusions(column, data))

        return conclusions

    def analyze_column(self, data):
        """
        Анализирует отдельный столбец данных.

        :param data: серия данных для анализа
        :return: словарь с результатами анализа
        """
        results = {}
        _, p_value_normal = stats.normaltest(data)
        results['normal_p_value'] = p_value_normal
        _, p_value_uniform = stats.kstest(data, 'uniform', args=(data.min(), data.max()))
        results['uniform_p_value'] = p_value_uniform
        median = data.median()
        binary = (data > median).astype(int)
        _, p_value_run = stats.mannwhitneyu(binary, 1 - binary, alternative='two-sided')
        results['run_test_p_value'] = p_value_run
        expected_model = self.fit_expected_model(data)
        deviation = self.calculate_deviation(data, expected_model)
        results['model_deviation'] = deviation
        return results

    def fit_expected_model(self, data):
        """
        Подгоняет линейную модель к данным.

        :param data: серия данных для подгонки модели
        :return: ожидаемые значения модели
        """
        x = np.arange(len(data))

        def linear_model(x, a, b):
            return a * x + b

        popt, _ = curve_fit(linear_model, x, data)
        return linear_model(x, *popt)

    def calculate_deviation(self, data, expected):
        """
        Вычисляет среднеквадратичное отклонение данных от ожидаемой модели.

        :param data: фактические данные
        :param expected: ожидаемые значения модели
        :return: среднеквадратичное отклонение
        """
        return np.sqrt(np.mean((data - expected) ** 2))

    async def generate_column_conclusions(self, column, data):
        conclusions = [f"\nАнализ столбца: {column}"]

        # Базовая статистика
        basic_stats = await asyncio.get_event_loop().run_in_executor(
            self.executor,
            lambda: {
                'mean': np.mean(data),
                'median': np.median(data),
                'std': np.std(data),
                'min': np.min(data),
                'max': np.max(data)
            }
        )

        conclusions.append("Базовая статистика:")
        conclusions.append(f"  Среднее: {basic_stats['mean']:.2f}")
        conclusions.append(f"  Медиана: {basic_stats['median']:.2f}")
        conclusions.append(f"  Стандартное отклонение: {basic_stats['std']:.2f}")
        conclusions.append(f"  Минимум: {basic_stats['min']:.2f}")
        conclusions.append(f"  Максимум: {basic_stats['max']:.2f}")

        # Анализ распределения
        _, p_value_normal = stats.normaltest(data)
        conclusions.append(f"\nАнализ распределения:")
        if p_value_normal < 0.05:
            conclusions.append("  Распределение значимо отличается от нормального (p < 0.05)")
        else:
            conclusions.append("  Распределение не отличается значимо от нормального (p >= 0.05)")

        # Анализ выбросов
        Q1 = np.percentile(data, 25)
        Q3 = np.percentile(data, 75)
        IQR = Q3 - Q1
        lower_bound = Q1 - 1.5 * IQR
        upper_bound = Q3 + 1.5 * IQR
        outliers = data[(data < lower_bound) | (data > upper_bound)]

        conclusions.append("\nАнализ выбросов:")
        if len(outliers) > 0:
            conclusions.append(f"  Обнаружено {len(outliers)} выбросов")
            conclusions.append(f"  Нижняя граница нормы: {lower_bound:.2f}")
            conclusions.append(f"  Верхняя граница нормы: {upper_bound:.2f}")
            conclusions.append(f"  Минимальный выброс: {np.min(outliers):.2f}")
            conclusions.append(f"  Максимальный выброс: {np.max(outliers):.2f}")
        else:
            conclusions.append("  Выбросов не обнаружено")

        # Анализ трендов
        trend_analysis = await asyncio.get_event_loop().run_in_executor(
            self.executor, self.analyze_trend, data
        )
        conclusions.append("\nАнализ трендов:")
        conclusions.append(f"  {trend_analysis}")

        # Анализ периодичности
        periodicity = await asyncio.get_event_loop().run_in_executor(
            self.executor, self.analyze_periodicity, data
        )
        conclusions.append("\nАнализ периодичности:")
        conclusions.append(f"  {periodicity}")

        return conclusions

    def analyze_trend(self, data):
        """Анализирует тренд в данных."""
        x = np.arange(len(data))
        slope, _, _, p_value, _ = stats.linregress(x, data)

        if p_value < 0.05:
            if slope > 0:
                return f"Обнаружен значимый возрастающий тренд (наклон: {slope:.4f}, p-value: {p_value:.4f})"
            else:
                return f"Обнаружен значимый убывающий тренд (наклон: {slope:.4f}, p-value: {p_value:.4f})"
        else:
            return "Значимого тренда не обнаружено"

    def analyze_periodicity(self, data):
        """Анализирует периодичность в данных."""
        autocorr = [pd.Series(data).autocorr(lag=i) for i in range(1, min(len(data) // 2, 50))]
        peaks = np.where((autocorr[1:-1] > autocorr[:-2]) & (autocorr[1:-1] > autocorr[2:]))[0] + 1

        if len(peaks) > 0:
            main_period = peaks[0] + 1
            return f"Обнаружена возможная периодичность с периодом {main_period}"
        else:
            return "Явной периодичности не обнаружено"

        return conclusions

    def safe_ratio(self, data):
        """
        Безопасно вычисляет отношение между последовательными элементами,
        избегая деления на ноль.

        :param data: массив данных
        :return: массив отношений
        """
        diff = np.diff(data)
        return np.divide(diff, data[:-1], out=np.zeros_like(diff), where=data[:-1] != 0)

    def get_outliers_info(self, data):
        """
        Определяет выбросы в данных, используя метод межквартильного размаха.

        :param data: серия данных для анализа
        :return: словарь с информацией о выбросах
        """
        Q1 = data.quantile(0.25)
        Q3 = data.quantile(0.75)
        IQR = Q3 - Q1
        lower_bound = Q1 - 1.5 * IQR
        upper_bound = Q3 + 1.5 * IQR
        outliers = data[(data < lower_bound) | (data > upper_bound)]
        return {
            'outliers': outliers,
            'min_outlier': outliers.min() if not outliers.empty else None,
            'max_outlier': outliers.max() if not outliers.empty else None
        }
