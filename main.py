import pandas as pd

# Чтение данных из Excel
df = pd.read_excel("data.xlsx")

# Удаление пробелов в начале и в конце строк для столбцов "Источник" и "Магистраль"
df['Источник'] = df['Источник'].str.strip()
df['Магистраль'] = df['Магистраль'].str.strip()

# Создание нового столбца с комбинацией "Источник+Магистраль"
df['Источник+Магистраль'] = df['Источник'] + '+' + df['Магистраль']

# Преобразование столбца "Дата" в формат datetime
df['Дата'] = pd.to_datetime(df['Дата'], errors='coerce')

# Извлечение уникальных дат (только даты)
unique_dates = sorted(df['Дата'].dropna().dt.date.unique())

# Извлечение уникальных комбинаций "Источник+Магистраль"
unique_combinations = df['Источник+Магистраль'].dropna().unique()

# Создание пустых DataFrame для:
# 1. Номеров записей
# 2. Температуры наружного воздуха (Тнв)
# 3. Фактической температуры наружного воздуха (Тнвф)
work_days_df = pd.DataFrame(index=unique_dates, columns=unique_combinations)
outer_temperature_df = pd.DataFrame(index=unique_dates, columns=unique_combinations)
actual_temperature_df = pd.DataFrame(index=unique_dates, columns=unique_combinations)

# Заполнение DataFrame-ов
for date in unique_dates:
    # Фильтрация записей по дате
    filtered_by_date = df[df['Дата'].dt.date == date]
    for combination in unique_combinations:
        # Фильтрация по комбинации "Источник+Магистраль"
        filtered_by_combination = filtered_by_date[filtered_by_date['Источник+Магистраль'] == combination]
        if not filtered_by_combination.empty:
            # Запись номера записи
            work_days_df.at[date, combination] = filtered_by_combination.iloc[0]['№№']
            # Запись температуры наружного воздуха (из столбца с именем "Тнв")
            outer_temperature_df.at[date, combination] = filtered_by_combination.iloc[0]['Тнв']
            # Запись фактической температуры наружного воздуха (Тнвф)
            # Данные для Тнвф находятся в 14-ом столбце исходного Excel файла (индекс 13)
            actual_temperature_df.at[date, combination] = filtered_by_combination.iloc[0, 13]

#############################################
# Запись результатов в Excel с выделением выбросов
#############################################
# Используем ExcelWriter с движком xlsxwriter (убедитесь, что установлен: pip install XlsxWriter)
with pd.ExcelWriter("output.xlsx", engine='xlsxwriter') as writer:
    # Лист с номерами записей
    work_days_df.to_excel(writer, sheet_name="Номера записей")
    # Лист с температурой наружного воздуха (Тнв)
    outer_temperature_df.to_excel(writer, sheet_name="Тнв")
    # Новый лист с фактической температурой наружного воздуха (Тнвф)
    actual_temperature_df.to_excel(writer, sheet_name="Тнвф")
    
    # Получаем объекты книги и листов для дальнейшего форматирования
    workbook = writer.book
    worksheet_tnv = writer.sheets["Тнв"]
    worksheet_tnvf = writer.sheets["Тнвф"]
    
    # Формат для выделения выбросов (красная заливка)
    red_format = workbook.add_format({'bg_color': '#FFC7CE'})
    
    # Функция для анализа выбросов по строкам (метод 2σ) и выделения их красной заливкой
    def highlight_outliers(worksheet, df_sheet, start_row=1, start_col=1):
        """
        Проходит по строкам DataFrame и выделяет ячейки с выбросами.
        Параметры start_row и start_col задают смещение (так как первая строка – заголовок, первый столбец – индекс).
        """
        for i, date in enumerate(df_sheet.index):
            row_series = df_sheet.loc[date]
            # Приводим значения к числовому типу (нечисловые значения станут NaN)
            row_values = pd.to_numeric(row_series, errors='coerce')
            valid = row_values.dropna()
            if len(valid) < 2:
                continue  # Недостаточно данных для расчёта среднего и стандартного отклонения
            row_mean = valid.mean()
            row_std = valid.std()
            if row_std == 0:
                continue  # Если стандартное отклонение равно 0, выбросов не обнаруживаем
            # Проходим по каждому столбцу строки
            for j, col in enumerate(df_sheet.columns):
                cell_value = row_series[col]
                try:
                    value = float(cell_value)
                except (ValueError, TypeError):
                    continue
                if abs(value - row_mean) > 2 * row_std:
                    # Записываем значение с красной заливкой
                    worksheet.write(i + start_row, j + start_col, value, red_format)
    
    # Применяем выделение выбросов для листа "Тнв"
    highlight_outliers(worksheet_tnv, outer_temperature_df)
    # Применяем выделение выбросов для листа "Тнвф"
    highlight_outliers(worksheet_tnvf, actual_temperature_df)
    
    # По завершении блока with данные сохраняются в output.xlsx

import seaborn as sns
import matplotlib.pyplot as plt

# Приводим данные к числовому типу
outer_numeric = outer_temperature_df.apply(pd.to_numeric, errors='coerce')
actual_numeric = actual_temperature_df.apply(pd.to_numeric, errors='coerce')

# 1. Линейный график: вычисляем по датам среднее значение по всем комбинациям
daily_predicted = outer_numeric.mean(axis=1)
daily_actual = actual_numeric.mean(axis=1)

# Собираем данные в один DataFrame для удобства построения графика
comparison_df = pd.DataFrame({
    'Тнв (метеорологи)': daily_predicted,
    'Тнвф (фактическая)': daily_actual
})
# Если индекс – дата в виде строки, приводим его к типу datetime:
comparison_df.index = pd.to_datetime(comparison_df.index)

# Сброс индекса для seaborn
comparison_reset = comparison_df.reset_index().rename(columns={'index': 'Дата'})

# Строим линейный график
plt.figure(figsize=(10, 6))
sns.lineplot(data=comparison_reset, x='Дата', y='Тнв (метеорологи)', marker='o', label='Метеорологи (Тнв)')
sns.lineplot(data=comparison_reset, x='Дата', y='Тнвф (фактическая)', marker='o', label='Фактическая (Тнвф)')
plt.title("Сравнение дневной средней температуры: метеорологи vs фактическая")
plt.xlabel("Дата")
plt.ylabel("Температура (°C)")
plt.legend()
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

# 2. Scatter plot: сравним значения Тнв и Тнвф для всех дат и комбинаций
# Преобразуем таблицы в формат «длинный» (stacked)
stacked_predicted = outer_numeric.stack().reset_index(name='Тнв')
stacked_actual = actual_numeric.stack().reset_index(name='Тнвф')

# Объединяем по дате и комбинации
merged = pd.merge(stacked_predicted, stacked_actual, on=['level_0', 'level_1'])
merged = merged.rename(columns={'level_0': 'Дата', 'level_1': 'Комбинация'})

plt.figure(figsize=(8, 6))
sns.scatterplot(data=merged, x='Тнв', y='Тнвф')
plt.title("Сравнение по точкам: метеорологи (Тнв) vs фактическая (Тнвф)")
plt.xlabel("Метеорологи (Тнв)")
plt.ylabel("Фактическая (Тнвф)")
# Добавляем линию y=x для визуальной оценки соответствия
min_val = min(merged['Тнв'].min(), merged['Тнвф'].min())
max_val = max(merged['Тнв'].max(), merged['Тнвф'].max())
plt.plot([min_val, max_val], [min_val, max_val], 'r--')
plt.tight_layout()
plt.savefig("plot.png")

