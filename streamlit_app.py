import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
import os
import xlsxwriter

# --- Настройки страницы ---
st.set_page_config(page_title="Программа анализа километров", layout="wide")

# --- Константы ---
DB_FILENAME = "database.parquet"
COLOR_MAP = {
    2: ('#FF0000', '#FFFFFF'),  # красный фон, белый текст
    3: ('#FFFF00', '#000000'),  # желтый фон, черный текст
    4: ('#ADD8E6', '#000000'),  # светло-голубой фон, черный текст
    5: ('#90EE90', '#000000')   # светло-зеленый фон, черный текст
}

# --- Функции ---

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

        cols = [
            "Дата", "ГОД", "МЕСЯЦ", "Вид проверки", "KM", "ПУТЬ", "КОДНАПР",
            "ПЧ", "ПД", "ОЦЕНКА", "БАЛЛ", "ПРОВЕРЕНО"
        ]
        if filename:
            cols.append("Файл")

        return df_filtered[cols].reset_index(drop=True)

    except Exception as e:
        st.error(f"Ошибка обработки файла {filename}: {e}")
        return pd.DataFrame()

def highlight_cells(val):
    if pd.isna(val):
        return "border: 1px solid black;"
    try:
        val_int = int(round(float(val)))
    except:
        return "border: 1px solid black;"

    if val_int in COLOR_MAP:
        bg, fg = COLOR_MAP[val_int]
        return f'background-color: {bg}; color: {fg}; border: 1px solid black;'
    return "border: 1px solid black;"

def save_styled_pivot_to_excel(pivot_table, filename):
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet("Сводная")

    formats = {
        2: workbook.add_format({'bg_color': COLOR_MAP[2][0], 'font_color': COLOR_MAP[2][1], 'border': 1}),
        3: workbook.add_format({'bg_color': COLOR_MAP[3][0], 'font_color': COLOR_MAP[3][1], 'border': 1}),
        4: workbook.add_format({'bg_color': COLOR_MAP[4][0], 'font_color': COLOR_MAP[4][1], 'border': 1}),
        5: workbook.add_format({'bg_color': COLOR_MAP[5][0], 'font_color': COLOR_MAP[5][1], 'border': 1}),
        'default': workbook.add_format({'border': 1})
    }

    # Заголовки
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
                    val_int = int(round(float(val)))
                    fmt = formats.get(val_int, formats['default'])
                    worksheet.write(row_idx, col_idx, val_int, fmt)
                except:
                    worksheet.write(row_idx, col_idx, val, formats['default'])

    workbook.close()

def render_legend():
    st.markdown("**Легенда цветов оценок:**")
    cols = st.columns(len(COLOR_MAP))
    for i, (score, (bg, fg)) in enumerate(COLOR_MAP.items()):
        with cols[i]:
            st.markdown(
                f'<div style="background-color:{bg};color:{fg};padding:10px;border-radius:5px;text-align:center;">{score}</div>',
                unsafe_allow_html=True)

# --- Интерфейс ---

# Тема
theme = st.sidebar.selectbox("🎨 Выберите тему", options=["Светлая", "Тёмная"])
if theme == "Тёмная":
    st.markdown(
        """
        <style>
            .main {background-color: #0E1117; color: white;}
            .css-1d391kg, .css-ffhzg2 {color: white;}
            .stButton>button {background-color: #333; color: white;}
            .stDataFrame div {color: white;}
        </style>
        """, unsafe_allow_html=True
    )
else:
    st.markdown(
        """
        <style>
            .main {background-color: white; color: black;}
        </style>
        """, unsafe_allow_html=True
    )

st.title("📊 Программа анализа километров")

# --- Загрузка данных ---
if os.path.exists(DB_FILENAME):
    base_df = pd.read_parquet(DB_FILENAME)
else:
    base_df = pd.DataFrame()

st.sidebar.header("📂 Загрузка файлов")
uploaded_files = st.sidebar.file_uploader(
    "Выберите Excel-файлы (.xlsx)",
    type="xlsx",
    accept_multiple_files=True
)

if uploaded_files:
    all_data = []
    upload_errors = []
    for file in uploaded_files:
        try:
            df_raw = pd.read_excel(file, sheet_name="Оценка КМ")
            processed = process_data(df_raw, file.name)
            if not processed.empty:
                all_data.append(processed)
            else:
                upload_errors.append(file.name)
        except Exception as e:
            upload_errors.append(f"{file.name}: {e}")

    if all_data:
        new_data = pd.concat(all_data, ignore_index=True)
        base_df = pd.concat([base_df, new_data], ignore_index=True).drop_duplicates()
        base_df.to_parquet(DB_FILENAME, index=False)
        st.sidebar.success(f"✅ Добавлено файлов: {len(all_data)}")
    if upload_errors:
        st.sidebar.error(f"Ошибки при загрузке: {upload_errors}")

if base_df.empty:
    st.info("📂 Загрузите Excel-файлы для начала работы.")
    st.stop()

# --- Сайдбар: фильтры ---
st.sidebar.header("⚙️ Фильтры")

# Фильтр по дате
min_date = base_df["Дата"].min()
max_date = base_df["Дата"].max()
date_range = st.sidebar.date_input(
    "Диапазон дат",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date
)
df_filtered = base_df.copy()
if isinstance(date_range, tuple) and len(date_range) == 2:
    df_filtered = df_filtered[
        (df_filtered["Дата"] >= pd.to_datetime(date_range[0])) &
        (df_filtered["Дата"] <= pd.to_datetime(date_range[1]))
    ]

# Множественный выбор км и путей
km_options = sorted(df_filtered["KM"].unique())
selected_kms = st.sidebar.multiselect("Выберите километры", km_options, default=km_options)

put_options = sorted(df_filtered["ПУТЬ"].unique())
selected_puts = st.sidebar.multiselect("Выберите пути", put_options, default=put_options)

df_filtered = df_filtered[
    (df_filtered["KM"].isin(selected_kms)) &
    (df_filtered["ПУТЬ"].isin(selected_puts))
]

# Метрика для анализа
metric = st.sidebar.selectbox("Метрика", ["ОЦЕНКА", "БАЛЛ"])

# Среднее/медиана
agg_func = st.sidebar.radio("Агрегатная функция для сводной таблицы", ["Среднее", "Медиана"])

# Кнопка удаления базы
if st.sidebar.button("🗑️ Удалить базу данных"):
    if st.sidebar.checkbox("Подтверждаю удаление базы"):
        if os.path.exists(DB_FILENAME):
            os.remove(DB_FILENAME)
            st.sidebar.warning("База удалена. Перезагрузите страницу.")
            st.stop()

# --- Статистика по фильтрованным данным ---
st.subheader("📊 Статистика данных")
st.markdown(f"- Количество записей: **{len(df_filtered):,}**")
st.markdown(f"- Даты: **{df_filtered['Дата'].min().date()}** — **{df_filtered['Дата'].max().date()}**")
st.markdown(f"- Уникальных километров: **{df_filtered['KM'].nunique()}**")
st.markdown(f"- Уникальных путей: **{df_filtered['ПУТЬ'].nunique()}**")
st.markdown(f"- Уникальных видов проверок: **{df_filtered['Вид проверки'].nunique()}**")

# --- График ---
st.subheader("📈 График по выбранным километрам и путям")

if df_filtered.empty:
    st.warning("Нет данных для выбранных фильтров.")
else:
    # Объединяем выбранные км и пути в фильтр
    fig_data = df_filtered.sort_values("Дата")

    # Построим линию с мультивыбором КМ и ПУТЬ
    fig = px.line(
        fig_data,
        x="Дата",
        y=metric,
        color='KM',
        line_dash='ПУТЬ',
        markers=True,
        title=f"{metric} по км и пути",
        labels={
            "Дата": "Дата",
            metric: metric,
            "KM": "Километр",
            "ПУТЬ": "Путь"
        },
        hover_data=["Вид проверки"]
    )
    fig.update_layout(legend_title_text='Километр / Путь')
    st.plotly_chart(fig, use_container_width=True)

# --- Таблица с pivot ---
st.subheader("📋 Сводная таблица")

# Добавим колонку для группировки по месяц-году и виду проверки
short_map = {"контрольная": "к", "рабочая": "р", "дополнительная": "д"}
df_filtered["МГ_Вид"] = df_filtered.apply(
    lambda row: f"{row['МЕСЯЦ']:02d}_{row['Дата'].year}_{short_map.get(row['Вид проверки'], '')}", axis=1
)

pivot = df_filtered.pivot_table(
    index=["KM", "ПУТЬ"],
    columns="МГ_Вид",
    values=metric,
    aggfunc="mean" if agg_func == "Среднее" else "median"
)

type_order = {'р': 0, 'к': 1, 'д': 2}
sorted_cols = sorted(
    pivot.columns,
    key=lambda x: (
        int(x.split('_')[1]),     # Год
        int(x.split('_')[0]),     # Месяц
        type_order.get(x.split('_')[2], 99)  # Тип проверки
    )
)
pivot = pivot[sorted_cols]

# Добавим столбец с агрегатом по выбранной функции
agg_series = None
if agg_func == "Среднее":
    agg_series = df_filtered.groupby(["KM", "ПУТЬ"])[metric].mean().round(2)
else:
    agg_series = df_filtered.groupby(["KM", "ПУТЬ"])[metric].median().round(2)

pivot["Итог"] = agg_series

# Округлим, используем тип Int64 для оценок
pivot = pivot.round(0).astype('Int64')

# Подсветка по оценкам, только если метрика — ОЦЕНКА
if metric == "ОЦЕНКА":
    styled_pivot = pivot.style.applymap(highlight_cells).format(lambda val: f"{val:.2f}" if isinstance(val, float) else val)
else:
    styled_pivot = pivot.style.format(lambda val: f"{val:.2f}" if isinstance(val, float) else val)

st.dataframe(styled_pivot, use_container_width=True, height=450)

# Легенда цветов
render_legend()

# --- Экспорт ---
st.subheader("📥 Скачать сводную таблицу")

output = BytesIO()
save_styled_pivot_to_excel(pivot, "styled_output.xlsx")

with open("styled_output.xlsx", "rb") as f:
    st.download_button(
        label="💾 Скачать Excel с подсветкой",
        data=f.read(),
        file_name="итоговая_таблица.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
