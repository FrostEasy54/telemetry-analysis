import base64
import io

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="Обработка Excel и визуализация", layout="wide")


def create_html_link(html_content, link_text="Открыть график"):
    b64 = base64.b64encode(html_content.encode("utf-8")).decode("utf-8")
    href = f'<a href="data:text/html;base64,{b64}" target="_blank" rel="noopener noreferrer">{link_text}</a>'
    return href


def process_data(df):
    color_sequence=[
        "#0068c9",
        "#ff2b2b",
        "#83c9ff",
        "#ffabab",
        "#29b09d",
        "#7defa1",
        "#ff8700",
        "#ffd16a",
        "#6d3fc0",
        "#d5dae5",
        "#f3dfc6",
        "#556b2f",
    ]
    # 1. Предварительная обработка данных
    df["Источник"] = df["Источник"].str.strip()
    df["Магистраль"] = df["Магистраль"].str.strip()
    df["Источник+Магистраль"] = df["Источник"] + "+" + df["Магистраль"]
    df["Дата"] = pd.to_datetime(df["Дата"], errors="coerce")
    unique_dates = sorted(df["Дата"].dropna().dt.date.unique())
    unique_combinations = df["Источник+Магистраль"].dropna().unique()

    # 2. Создание пустых DataFrame для различных наборов данных
    work_days_df = pd.DataFrame(index=unique_dates, columns=unique_combinations)
    outer_temperature_df = pd.DataFrame(index=unique_dates, columns=unique_combinations)
    actual_temperature_df = pd.DataFrame(
        index=unique_dates, columns=unique_combinations
    )

    water_direct_pred_df = pd.DataFrame(
        index=unique_dates, columns=unique_combinations
    )  # прямая (график)
    water_reverse_pred_df = pd.DataFrame(
        index=unique_dates, columns=unique_combinations
    )  # обратная (график)
    water_direct_act_df = pd.DataFrame(
        index=unique_dates, columns=unique_combinations
    )  # прямая (факт)
    water_reverse_act_df = pd.DataFrame(
        index=unique_dates, columns=unique_combinations
    )  # обратная (факт)

    pressure_df = pd.DataFrame(index=unique_dates)
    for combination in unique_combinations:
        pressure_df[f"{combination} - прямое давление"] = None
        pressure_df[f"{combination} - обратное давление"] = None

    flow_df = pd.DataFrame(index=unique_dates)
    for combination in unique_combinations:
        flow_df[f"{combination} - прямой расход"] = None
        flow_df[f"{combination} - обратный расход"] = None

    # 3. Заполнение DataFrame данными
    for date in unique_dates:
        filtered_by_date = df[df["Дата"].dt.date == date]
        for combination in unique_combinations:
            filtered_by_combination = filtered_by_date[
                filtered_by_date["Источник+Магистраль"] == combination
            ]
            if not filtered_by_combination.empty:
                # Номера записей и температуры наружного воздуха
                work_days_df.at[date, combination] = filtered_by_combination.iloc[0][
                    "№№"
                ]
                outer_temperature_df.at[date, combination] = (
                    filtered_by_combination.iloc[0]["Тнв"]
                )
                actual_temperature_df.at[date, combination] = (
                    filtered_by_combination.iloc[0, 13]
                )  # фактическая Тнв (14-ый столбец, индекс 13)

                # Температура теплоносителя (воды)
                water_direct_pred_df.at[date, combination] = (
                    filtered_by_combination.iloc[0, 4]
                )  # 5-ый столбец: прямая (график)
                water_reverse_pred_df.at[date, combination] = (
                    filtered_by_combination.iloc[0, 5]
                )  # 6-ой столбец: обратная (график)
                water_direct_act_df.at[date, combination] = (
                    filtered_by_combination.iloc[0, 6]
                )  # 7-ой столбец: прямая (факт)
                water_reverse_act_df.at[date, combination] = (
                    filtered_by_combination.iloc[0, 7]
                )  # 8-ой столбец: обратная (факт)

                # Давление: прямое – 15-ый столбец (индекс 14), обратное – 16-ый столбец (индекс 15)
                pressure_df.at[date, f"{combination} - прямое давление"] = (
                    filtered_by_combination.iloc[0, 14]
                )
                pressure_df.at[date, f"{combination} - обратное давление"] = (
                    filtered_by_combination.iloc[0, 15]
                )

                # Расход: прямой – 10-ый столбец (индекс 9), обратный – 11-ый столбец (индекс 10)
                flow_df.at[date, f"{combination} - прямой расход"] = (
                    filtered_by_combination.iloc[0, 9]
                )
                flow_df.at[date, f"{combination} - обратный расход"] = (
                    filtered_by_combination.iloc[0, 10]
                )

    # 4. Создание комбинированных DataFrame для температуры теплоносителя
    water_pred_combined_df = pd.DataFrame(index=unique_dates)
    water_act_combined_df = pd.DataFrame(index=unique_dates)
    for combination in unique_combinations:
        water_pred_combined_df[f"{combination} - прямая температура"] = (
            water_direct_pred_df[combination]
        )
        water_pred_combined_df[f"{combination} - обратная температура"] = (
            water_reverse_pred_df[combination]
        )
        water_act_combined_df[f"{combination} - прямая температура"] = (
            water_direct_act_df[combination]
        )
        water_act_combined_df[f"{combination} - обратная температура"] = (
            water_reverse_act_df[combination]
        )

    # 5. Запись результатов в Excel с форматированием
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
        # Запись листов
        work_days_df.to_excel(writer, sheet_name="Номера записей")
        outer_temperature_df.to_excel(writer, sheet_name="Тнв")
        actual_temperature_df.to_excel(writer, sheet_name="Тнвф")
        water_direct_pred_df.to_excel(
            writer, sheet_name="теплоноситель прямой (график)"
        )
        water_reverse_pred_df.to_excel(
            writer, sheet_name="теплоноситель обратный (график)"
        )
        water_direct_act_df.to_excel(writer, sheet_name="теплоноситель прямой (факт)")
        water_reverse_act_df.to_excel(
            writer, sheet_name="теплоноситель обратный (факт)"
        )
        pressure_df.to_excel(writer, sheet_name="Давление", startrow=2, header=False)
        water_pred_combined_df.to_excel(
            writer, sheet_name="Теплоноситель (график)", startrow=2, header=False
        )
        water_act_combined_df.to_excel(
            writer, sheet_name="Теплоноситель (факт)", startrow=2, header=False
        )
        flow_df.to_excel(writer, sheet_name="Расход", startrow=2, header=False)

        workbook = writer.book
        # Получаем объекты листов
        worksheet_tnv = writer.sheets["Тнв"]
        worksheet_tnvf = writer.sheets["Тнвф"]
        worksheet_water_direct_pred = writer.sheets["теплоноситель прямой (график)"]
        worksheet_water_reverse_pred = writer.sheets["теплоноситель обратный (график)"]
        worksheet_water_direct_act = writer.sheets["теплоноситель прямой (факт)"]
        worksheet_water_reverse_act = writer.sheets["теплоноситель обратный (факт)"]
        worksheet_pressure = writer.sheets["Давление"]
        worksheet_water_pred = writer.sheets["Теплоноситель (график)"]
        worksheet_water_act = writer.sheets["Теплоноситель (факт)"]
        worksheet_flow = writer.sheets["Расход"]

        # Определяем формат заголовков
        header_format = workbook.add_format(  # type: ignore
            { 
                "bold": True,
                "align": "center",
                "valign": "vcenter",
                "bg_color": "#D7E4BC",
                "border": 1,
            }
        )

        # Объединённые заголовки для листа "Давление"
        worksheet_pressure.merge_range(0, 0, 1, 0, "Дата", header_format)
        for i, combination in enumerate(unique_combinations):
            col_start = 1 + i * 2
            col_end = col_start + 1
            worksheet_pressure.merge_range(
                0, col_start, 0, col_end, combination, header_format
            )
            worksheet_pressure.write(1, col_start, "прямое давление", header_format)
            worksheet_pressure.write(1, col_end, "обратное давление", header_format)

        # Объединённые заголовки для листа "Расход"
        worksheet_flow.merge_range(0, 0, 1, 0, "Дата", header_format)
        for i, combination in enumerate(unique_combinations):
            col_start = 1 + i * 2
            col_end = col_start + 1
            worksheet_flow.merge_range(
                0, col_start, 0, col_end, combination, header_format
            )
            worksheet_flow.write(1, col_start, "прямой расход", header_format)
            worksheet_flow.write(1, col_end, "обратный расход", header_format)

        # Объединённые заголовки для листа "Теплоноситель (график)"
        worksheet_water_pred.merge_range(0, 0, 1, 0, "Дата", header_format)
        for i, combination in enumerate(unique_combinations):
            col_start = 1 + i * 2
            col_end = col_start + 1
            worksheet_water_pred.merge_range(
                0, col_start, 0, col_end, combination, header_format
            )
            worksheet_water_pred.write(
                1, col_start, "прямая температура", header_format
            )
            worksheet_water_pred.write(
                1, col_end, "обратная температура", header_format
            )

        # Объединённые заголовки для листа "Теплоноситель (факт)"
        worksheet_water_act.merge_range(0, 0, 1, 0, "Дата", header_format)
        for i, combination in enumerate(unique_combinations):
            col_start = 1 + i * 2
            col_end = col_start + 1
            worksheet_water_act.merge_range(
                0, col_start, 0, col_end, combination, header_format
            )
            worksheet_water_act.write(1, col_start, "прямая температура", header_format)
            worksheet_water_act.write(1, col_end, "обратная температура", header_format)

        # Формат для выделения выбросов (красная заливка)
        red_format = workbook.add_format({"bg_color": "#FFC7CE"})  # type: ignore

        # Условное форматирование для листа "Давление"
        for row_num, (date, row_data) in enumerate(pressure_df.iterrows()):
            excel_row = row_num + 2  # данные начинаются с 3-ей строки
            for i, combination in enumerate(unique_combinations):
                direct_col = f"{combination} - прямое давление"
                reverse_col = f"{combination} - обратное давление"
                try:
                    direct_value = float(row_data[direct_col])
                    reverse_value = float(row_data[reverse_col])
                except (ValueError, TypeError):
                    continue
                if reverse_value > direct_value:
                    excel_col = 1 + i * 2 + 1
                    worksheet_pressure.write(
                        excel_row, excel_col, reverse_value, red_format
                    )

        # Условное форматирование для листа "Теплоноситель (график)"
        for row_num, (date, row_data) in enumerate(water_pred_combined_df.iterrows()):
            excel_row = row_num + 2
            for i, combination in enumerate(unique_combinations):
                direct_col = f"{combination} - прямая температура"
                reverse_col = f"{combination} - обратная температура"
                try:
                    direct_value = float(row_data[direct_col])
                    reverse_value = float(row_data[reverse_col])
                except (ValueError, TypeError):
                    continue
                if reverse_value > direct_value:
                    excel_col = 1 + i * 2 + 1
                    worksheet_water_pred.write(
                        excel_row, excel_col, reverse_value, red_format
                    )

        # Условное форматирование для листа "Теплоноситель (факт)"
        for row_num, (date, row_data) in enumerate(water_act_combined_df.iterrows()):
            excel_row = row_num + 2
            for i, combination in enumerate(unique_combinations):
                direct_col = f"{combination} - прямая температура"
                reverse_col = f"{combination} - обратная температура"
                try:
                    direct_value = float(row_data[direct_col])
                    reverse_value = float(row_data[reverse_col])
                except (ValueError, TypeError):
                    continue
                if reverse_value > direct_value:
                    excel_col = 1 + i * 2 + 1
                    worksheet_water_act.write(
                        excel_row, excel_col, reverse_value, red_format
                    )

        # Условное форматирование для листа "Расход"
        for row_num, (date, row_data) in enumerate(flow_df.iterrows()):
            excel_row = row_num + 2
            for i, combination in enumerate(unique_combinations):
                try:
                    direct_value = float(row_data[f"{combination} - прямой расход"])
                    reverse_value = float(row_data[f"{combination} - обратный расход"])
                except (ValueError, TypeError):
                    continue
                if reverse_value > direct_value:
                    excel_col = 1 + i * 2 + 1
                    worksheet_flow.write(
                        excel_row, excel_col, reverse_value, red_format
                    )

        # Функция для выделения выбросов по строкам (метод 2σ)
        def highlight_outliers(worksheet, df_sheet, start_row=1, start_col=1):
            for i, date in enumerate(df_sheet.index):
                row_series = df_sheet.loc[date]
                row_values = pd.to_numeric(row_series, errors="coerce")
                valid = row_values.dropna()
                if len(valid) < 2:
                    continue
                row_mean = valid.mean()
                row_std = valid.std()
                if row_std == 0:
                    continue
                for j, col in enumerate(df_sheet.columns):
                    cell_value = row_series[col]
                    try:
                        value = float(cell_value)
                    except (ValueError, TypeError):
                        continue
                    if abs(value - row_mean) > 2 * row_std:
                        worksheet.write(i + start_row, j + start_col, value, red_format)

        # Применяем highlight_outliers к соответствующим листам
        highlight_outliers(worksheet_tnv, outer_temperature_df)
        highlight_outliers(worksheet_tnvf, actual_temperature_df)
        highlight_outliers(worksheet_water_direct_pred, water_direct_pred_df)
        highlight_outliers(worksheet_water_reverse_pred, water_reverse_pred_df)
        highlight_outliers(worksheet_water_direct_act, water_direct_act_df)
        highlight_outliers(worksheet_water_reverse_act, water_reverse_act_df)

    excel_data = excel_buffer.getvalue()

    # 6. Построение графиков с Plotly

    # Графики для наружного воздуха
    outer_numeric = outer_temperature_df.apply(pd.to_numeric, errors="coerce")
    actual_numeric = actual_temperature_df.apply(pd.to_numeric, errors="coerce")
    daily_predicted = outer_numeric.mean(axis=1)
    daily_actual = actual_numeric.mean(axis=1)
    comparison_df = pd.DataFrame(
        {
            "Дата": daily_predicted.index,
            "Тнв (метеорологи)": daily_predicted.values,
            "Тнвф (фактическая)": daily_actual.values,
        }
    )
    comparison_df["Дата"] = pd.to_datetime(comparison_df["Дата"])
    fig_line = px.line(
        comparison_df,
        x="Дата",
        y=["Тнв (метеорологи)", "Тнвф (фактическая)"],
        title="Сравнение средней температуры наружного воздуха",
        labels={"value": "Температура (°C)", "variable": "Тип данных"},
        color_discrete_sequence=color_sequence
    )

    stacked_predicted = outer_numeric.stack().reset_index(name="Тнв")  # type: ignore
    stacked_actual = actual_numeric.stack().reset_index(name="Тнвф")  # type: ignore
    merged = pd.merge(stacked_predicted, stacked_actual, on=["level_0", "level_1"])
    merged = merged.rename(columns={"level_0": "Дата", "level_1": "Комбинация"})
    merged["Дата"] = pd.to_datetime(merged["Дата"])
    fig_scatter = px.scatter(
        merged,
        x="Тнв",
        y="Тнвф",
        color="Комбинация",
        hover_data=["Дата", "Комбинация"],
        title="Сравнение по точкам: Тнв (метеорологи) vs Тнвф (факт)",
        labels={"Тнв": "Тнв (метеорологи)", "Тнвф": "Тнвф (фактическая)"},
        color_discrete_sequence=color_sequence
    )
    min_val = min(merged["Тнв"].min(), merged["Тнвф"].min())
    max_val = max(merged["Тнв"].max(), merged["Тнвф"].max())
    fig_scatter.add_trace(
        go.Scatter(
            x=[min_val, max_val],
            y=[min_val, max_val],
            mode="lines",
            line=dict(dash="dash", color="red"),
            name="y=x",
        )
    )

    # Графики для теплоносителя (прямая линия)
    water_direct_pred_numeric = water_direct_pred_df.apply(
        pd.to_numeric, errors="coerce"
    )
    water_direct_act_numeric = water_direct_act_df.apply(pd.to_numeric, errors="coerce")
    daily_water_direct_pred = water_direct_pred_numeric.mean(axis=1)
    daily_water_direct_act = water_direct_act_numeric.mean(axis=1)
    comparison_water_direct_df = pd.DataFrame(
        {
            "Дата": daily_water_direct_pred.index,
            "теплоноситель прямой (график)": daily_water_direct_pred.values,
            "теплоноситель прямой (факт)": daily_water_direct_act.values,
        }
    )
    comparison_water_direct_df["Дата"] = pd.to_datetime(
        comparison_water_direct_df["Дата"]
    )
    fig_line_water_direct = px.line(
        comparison_water_direct_df,
        x="Дата",
        y=["теплоноситель прямой (график)", "теплоноситель прямой (факт)"],
        title="Температура теплоносителя (прямая): график vs факт",
        labels={"value": "Температура (°C)", "variable": "Тип данных"},
        color_discrete_sequence=color_sequence
    )

    stacked_water_direct_pred = water_direct_pred_numeric.stack().reset_index(
        name="теплоноситель прямой (график)"
    )  # type: ignore
    stacked_water_direct_act = water_direct_act_numeric.stack().reset_index(
        name="теплоноситель прямой (факт)"
    )  # type: ignore
    merged_water_direct = pd.merge(
        stacked_water_direct_pred, stacked_water_direct_act, on=["level_0", "level_1"]
    )
    merged_water_direct = merged_water_direct.rename(
        columns={"level_0": "Дата", "level_1": "Комбинация"}
    )
    merged_water_direct["Дата"] = pd.to_datetime(merged_water_direct["Дата"])
    fig_scatter_water_direct = px.scatter(
        merged_water_direct,
        x="теплоноситель прямой (график)",
        y="теплоноситель прямой (факт)",
        color="Комбинация",
        hover_data=["Дата", "Комбинация"],
        title="Сравнение по точкам: теплоноситель прямой – график vs факт",
        labels={
            "теплоноситель прямой (график)": "график",
            "теплоноситель прямой (факт)": "факт",
        },
        color_discrete_sequence=color_sequence,
    )
    min_val_direct = min(
        merged_water_direct["теплоноситель прямой (график)"].min(),
        merged_water_direct["теплоноситель прямой (факт)"].min(),
    )
    max_val_direct = max(
        merged_water_direct["теплоноситель прямой (график)"].max(),
        merged_water_direct["теплоноситель прямой (факт)"].max(),
    )
    fig_scatter_water_direct.add_trace(
        go.Scatter(
            x=[min_val_direct, max_val_direct],
            y=[min_val_direct, max_val_direct],
            mode="lines",
            line=dict(dash="dash", color="red"),
            name="y=x",
        )
    )

    # Графики для теплоносителя (обратная линия)
    water_reverse_pred_numeric = water_reverse_pred_df.apply(
        pd.to_numeric, errors="coerce"
    )
    water_reverse_act_numeric = water_reverse_act_df.apply(
        pd.to_numeric, errors="coerce"
    )
    daily_water_reverse_pred = water_reverse_pred_numeric.mean(axis=1)
    daily_water_reverse_act = water_reverse_act_numeric.mean(axis=1)
    comparison_water_reverse_df = pd.DataFrame(
        {
            "Дата": daily_water_reverse_pred.index,
            "теплоноситель обратный (график)": daily_water_reverse_pred.values,
            "теплоноситель обратный (факт)": daily_water_reverse_act.values,
        }
    )
    comparison_water_reverse_df["Дата"] = pd.to_datetime(
        comparison_water_reverse_df["Дата"]
    )
    fig_line_water_reverse = px.line(
        comparison_water_reverse_df,
        x="Дата",
        y=["теплоноситель обратный (график)", "теплоноситель обратный (факт)"],
        title="Температура теплоносителя (обратный): график vs факт",
        labels={"value": "Температура (°C)", "variable": "Тип данных"},
        color_discrete_sequence=color_sequence,
    )

    stacked_water_reverse_pred = water_reverse_pred_numeric.stack().reset_index(
        name="теплоноситель обратный (график)"
    )  # type: ignore
    stacked_water_reverse_act = water_reverse_act_numeric.stack().reset_index(
        name="теплоноситель обратный (факт)"
    )  # type: ignore
    merged_water_reverse = pd.merge(
        stacked_water_reverse_pred, stacked_water_reverse_act, on=["level_0", "level_1"]
    )
    merged_water_reverse = merged_water_reverse.rename(
        columns={"level_0": "Дата", "level_1": "Комбинация"}
    )
    merged_water_reverse["Дата"] = pd.to_datetime(merged_water_reverse["Дата"])
    fig_scatter_water_reverse = px.scatter(
        merged_water_reverse,
        x="теплоноситель обратный (график)",
        y="теплоноситель обратный (факт)",
        color="Комбинация",
        hover_data=["Дата", "Комбинация"],
        title="Сравнение по точкам: теплоноситель обратный – график vs факт",
        labels={
            "теплоноситель обратный (график)": "график",
            "теплоноситель обратный (факт)": "факт",
        },
        color_discrete_sequence=color_sequence,
    )
    min_val_reverse = min(
        merged_water_reverse["теплоноситель обратный (график)"].min(),
        merged_water_reverse["теплоноситель обратный (факт)"].min(),
    )
    max_val_reverse = max(
        merged_water_reverse["теплоноситель обратный (график)"].max(),
        merged_water_reverse["теплоноситель обратный (факт)"].max(),
    )
    fig_scatter_water_reverse.add_trace(
        go.Scatter(
            x=[min_val_reverse, max_val_reverse],
            y=[min_val_reverse, max_val_reverse],
            mode="lines",
            line=dict(dash="dash", color="red"),
            name="y=x",
        )
    )

    # Новый график: комбинированный график для теплоносителя (прямая и обратная) и наружного воздуха (график)
    daily_water_direct = water_direct_pred_numeric.mean(axis=1)
    daily_water_reverse = water_reverse_pred_numeric.mean(axis=1)
    daily_outer = outer_numeric.mean(axis=1)
    combined_df = pd.DataFrame(
        {
            "Дата": daily_water_direct.index,
            "Прямая температура теплоносителя": daily_water_direct.values,
            "Обратная температура теплоносителя": daily_water_reverse.values,
            "Температура наружного воздуха": daily_outer.values,
        }
    )
    combined_df["Дата"] = pd.to_datetime(combined_df["Дата"])
    fig_new = px.line(
        combined_df,
        x="Дата",
        y=[
            "Прямая температура теплоносителя",
            "Обратная температура теплоносителя",
            "Температура наружного воздуха",
        ],
        title="Сравнение средней температуры теплоносителя и наружного воздуха (график)",
        labels={"value": "Температура (°C)", "variable": "Показатель"},
        color_discrete_sequence=color_sequence,
    )

    # Новый график: комбинированный график фактических данных теплоносителя и наружного воздуха
    water_direct_act_numeric = water_direct_act_df.apply(pd.to_numeric, errors="coerce")
    water_reverse_act_numeric = water_reverse_act_df.apply(
        pd.to_numeric, errors="coerce"
    )
    actual_numeric = actual_temperature_df.apply(pd.to_numeric, errors="coerce")
    daily_water_direct_act = water_direct_act_numeric.mean(axis=1)
    daily_water_reverse_act = water_reverse_act_numeric.mean(axis=1)
    daily_actual_outer = actual_numeric.mean(axis=1)
    combined_act_df = pd.DataFrame(
        {
            "Дата": daily_water_direct_act.index,
            "Прямая температура теплоносителя (факт)": daily_water_direct_act.values,
            "Обратная температура теплоносителя (факт)": daily_water_reverse_act.values,
            "Фактическая температура наружного воздуха": daily_actual_outer.values,
        }
    )
    combined_act_df["Дата"] = pd.to_datetime(combined_act_df["Дата"])
    fig_act = px.line(
        combined_act_df,
        x="Дата",
        y=[
            "Прямая температура теплоносителя (факт)",
            "Обратная температура теплоносителя (факт)",
            "Фактическая температура наружного воздуха",
        ],
        title="Сравнение средней фактической температуры теплоносителя и наружного воздуха",
        labels={"value": "Температура (°C)", "variable": "Показатель"},
        color_discrete_sequence=color_sequence,
    )

    # Собираем 8 графиков в словарь
    graphs = {
        "line_comparison_outer_temp.html": fig_line.to_html(),
        "scatter_comparison_outer_temp.html": fig_scatter.to_html(),
        "line_comparison_water_direct.html": fig_line_water_direct.to_html(),
        "scatter_comparison_water_direct.html": fig_scatter_water_direct.to_html(),
        "line_comparison_water_reverse.html": fig_line_water_reverse.to_html(),
        "scatter_comparison_water_reverse.html": fig_scatter_water_reverse.to_html(),
        "line_comparison_water_and_outer_temp.html": fig_new.to_html(),
        "line_comparison_water_and_outer_temp_actual.html": fig_act.to_html(),
    }

    return excel_data, graphs


def main():
    st.title("Обработка данных телеметрии")
    st.write("Загрузите ваш Excel с данными телеметрии")

    uploaded_file = st.file_uploader("Выберите Excel файл", type=["xlsx"])

    if st.button("Запустить обработку"):
        if uploaded_file is not None:
            try:
                df = pd.read_excel(uploaded_file)
            except Exception as e:
                st.error(f"Ошибка при чтении файла: {e}")
                return

            with st.spinner("Обработка данных..."):
                excel_data, graphs = process_data(df)
            st.success("Обработка завершена!")

            st.subheader("Скачать обработанный Excel файл:")
            st.download_button(
                "Скачать output.xlsx",
                data=excel_data,
                file_name="output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            # st.subheader("Скачать HTML графики")
            # for filename, html_content in graphs.items():
            #     st.download_button(
            #         f"Скачать {filename}",
            #         data=html_content,
            #         file_name=filename,
            #         mime="text/html",
            #     )
            st.subheader("Встроенные HTML графики:")
            with st.spinner("Встраивание графиков..."):
                # Встраиваем все графики сразу на страницу
                for filename, html_content in graphs.items():
                    st.markdown(f"#### {filename}")
                    components.html(html_content, height=600)
        else:
            st.error("Пожалуйста, загрузите Excel файл.")


if __name__ == "__main__":
    main()
