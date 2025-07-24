import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
import os
import xlsxwriter

st.set_page_config(page_title="Программа анализа километров", layout="wide")
st.title("\U0001F4CA Программа анализа километров")

DB_FILENAME = "database.parquet"  # Файл для хранения базы

# ====== ФУНКЦИИ ======

def process_data(df, filename=None):
    try:
        df_filtered = df[
            (df['ПЧ'] == 22) &
            (df['ПУТЬ'].isin([1, 2])) &
            (df['КОДНАПР'].isin([24602, 24701])) &
            (df['ПД'].isin([4, 5, 12])) &
            (((df['KM'] >= 103) & (df['KM'] <= 175)) | ((df['KM'] >= 2342) & (df['KM'] <= 2346))) &
            (df['ОЦЕНКА'].isin([2, 3, 4, 5]))
        ].copy()

        df_filtered["Дата"] = pd.to_datetime(dict(
            year=df_filtered["ГОД"],
            month=df_filtered["МЕСЯЦ"],
            day=df_filtered["ДЕНЬ"]
        ), errors='coerce')

        df_filtered["Вид проверки"] = df_filtered["ВИД"].map({
            0: "рабочая",
            1: "контрольная",
            2: "дополнительная"
        })

        if filename:
            df_filtered["Файл"] = filename

        df_filtered["ОЦЕНКА"] = df_filtered["ОЦЕНКА"].astype(int)

        return df_filtered[[
            "Дата", "ГОД", "МЕСЯЦ", "Вид проверки", "KM", "ПУТЬ", "КОДНАПР",
            "ПЧ", "ПД", "ОЦЕНКА", "БАЛЛ", "ПРОВЕРЕНО", "Файл" if filename else None
        ]].reset_index(drop=True)

    except Exception as e:
        st.error(f"Ошибка обработки файла {filename}: {e}")
        return pd.DataFrame()

def highlight_cells(val):
    if pd.isna(val):
        return "border: 1px solid black;"
    try:
        val = float(val)
        val_int = int(round(val))
    except:
        return "border: 1px solid black;"

    if val_int == 2:
        return 'background-color: red; color: white; border: 1px solid black;'
    elif val_int == 3:
        return 'background-color: yellow; color: black; border: 1px solid black;'
    elif val_int == 4:
        return 'background-color: lightblue; color: black; border: 1px solid black;'
    elif val_int == 5:
        return 'background-color: lightgreen; color: black; border: 1px solid black;'
    return "border: 1px solid black;"

def save_styled_pivot_to_excel(pivot_table, filename):
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet("Сводная")

    formats = {
        2: workbook.add_format({'bg_color': '#FF0000', 'font_color': '#FFFFFF', 'border': 1}),
        3: workbook.add_format({'bg_color': '#FFFF00', 'font_color': '#000000', 'border': 1}),
        4: workbook.add_format({'bg_color': '#ADD8E6', 'font_color': '#000000', 'border': 1}),
        5: workbook.add_format({'bg_color': '#90EE90', 'font_color': '#000000', 'border': 1}),
        'default': workbook.add_format({'border': 1})
    }

    worksheet.write(0, 0, "KM")
    worksheet.write(0, 1, "ПУТЬ")
    for col_idx, col in enumerate(pivot_table.columns, start=2):
        worksheet.write(0, col_idx, col)

    for row_idx, (index, row) in enumerate(pivot_table.iterrows(), start=1):
        worksheet.write(row_idx, 0, index[0])
        worksheet.write(row_idx, 1, index[1])
        for col_idx, val in enumerate(row, start=2):
            if pd.isna(val):
                worksheet.write(row_idx, col_idx, "", formats['default'])
            else:
                try:
                    val_int = float(val)
                    fmt = formats.get(int(round(val_int)), formats['default'])
                    worksheet.write(row_idx, col_idx, val_int, fmt)
                except:
                    worksheet.write(row_idx, col_idx, val, formats['default'])

    workbook.close()

# ====== ЗАГРУЗКА/ХРАНЕНИЕ БАЗЫ ======
if os.path.exists(DB_FILENAME):
    base_df = pd.read_parquet(DB_FILENAME)
else:
    base_df = pd.DataFrame()

uploaded_files = st.file_uploader("\U0001F4C2 Загрузите Excel-файлы", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for file in uploaded_files:
        df_raw = pd.read_excel(file, sheet_name="Оценка КМ")
        processed = process_data(df_raw, file.name)
        all_data.append(processed)

    if all_data:
        new_data = pd.concat(all_data, ignore_index=True)
        base_df = pd.concat([base_df, new_data], ignore_index=True).drop_duplicates()
        base_df.to_parquet(DB_FILENAME, index=False)
        st.success("✅ Данные добавлены в базу")

# ====== АНАЛИЗ ======
if not base_df.empty:
    st.subheader("\U0001F4C4 Предпросмотр данных")
    st.dataframe(base_df.head(100), use_container_width=True)

    if st.button("🗑️ Удалить базу"):
        os.remove(DB_FILENAME)
        st.warning("База удалена. Перезагрузите страницу.")
        st.stop()

    st.subheader("\U0001F4C5 Фильтр по дате")
    min_date = base_df["Дата"].min()
    max_date = base_df["Дата"].max()
    date_range = st.date_input("Выберите диапазон дат", (min_date, max_date))

    df_filtered = base_df.copy()
    if isinstance(date_range, tuple) and len(date_range) == 2:
        df_filtered = df_filtered[(df_filtered["Дата"] >= pd.to_datetime(date_range[0])) &
                                  (df_filtered["Дата"] <= pd.to_datetime(date_range[1]))]

    metric = st.selectbox("Выберите метрику", ["ОЦЕНКА", "БАЛЛ"])

    st.subheader("\U0001F4C8 График по км")
    km = st.selectbox("Километр", sorted(df_filtered["KM"].unique()))
    put = st.selectbox("Путь", sorted(df_filtered["ПУТЬ"].unique()))
    df_km = df_filtered[(df_filtered["KM"] == km) & (df_filtered["ПУТЬ"] == put)]

    if not df_km.empty:
        fig = px.line(df_km.sort_values("Дата"), x="Дата", y=metric, markers=True,
                      title=f"{metric} — км {km}, путь {put}")
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("\U0001F4CB Таблица по км, пути и проверкам")

    short_map = {"контрольная": "к", "рабочая": "р", "дополнительная": "д"}
    df_filtered["МГ_Вид"] = df_filtered.apply(
        lambda row: f"{row['МЕСЯЦ']:02d}_{row['Дата'].year}_{short_map.get(row['Вид проверки'], '')}", axis=1
    )

    pivot = df_filtered.pivot_table(
        index=["KM", "ПУТЬ"],
        columns="МГ_Вид",
        values=metric,
        aggfunc="mean"
    )

    pivot = pivot[sorted(pivot.columns, key=lambda x: (int(x.split('_')[1]), int(x.split('_')[0]), x.split('_')[2]))]
    pivot = pivot.round(0).astype('Int64')
    pivot["Среднее"] = df_filtered.groupby(["KM", "ПУТЬ"])[metric].mean().round(2)

    st.dataframe(
        pivot.style
            .format(lambda val: f"{val:.2f}" if isinstance(val, float) else val)
            .applymap(highlight_cells),
        use_container_width=True
    )

    st.subheader("\U0001F4E5 Скачать таблицу")
    output = BytesIO()
    save_styled_pivot_to_excel(pivot, "styled_output.xlsx")
    with open("styled_output.xlsx", "rb") as f:
        st.download_button(
            label="\U0001F4BE Скачать Excel с подсветкой",
            data=f.read(),
            file_name="итоговая_таблица.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("\U0001F4C2 Загрузите Excel-файлы для начала работы.")