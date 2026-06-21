import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(layout="wide")
st.title("📊 Биллинг Магнит")

# ===================== ЗАГРУЗКА =====================
url_main = "https://docs.google.com/spreadsheets/d/1-kv25KvN60XPksMm4YC27r_x9cxn0PevY4npWWdOi4c/export?format=csv&gid=348938437"
df = pd.read_csv(url_main)
df.columns = df.columns.str.strip()

# ===================== ПОИСК КОЛОНОК =====================
def find_col(name):
    for col in df.columns:
        if name.lower() in col.lower():
            return col
    return None

df = df.rename(columns={
    find_col('дата'): 'Дата',
    find_col('шт'): 'Кол-во шт',
    find_col('сумм'): 'Итого сумма',
    find_col('регион'): 'Регион',
    find_col('авто'): 'Тип ТС',
    find_col('магаз'): 'Магазин'
})

# ===================== ЧИСТКА =====================
df['Дата'] = pd.to_datetime(df['Дата'], errors='coerce')

df['Кол-во шт'] = pd.to_numeric(
    df['Кол-во шт'].astype(str).str.replace(" ", "").str.replace(",", "."),
    errors='coerce'
).fillna(0)

df['Итого сумма'] = pd.to_numeric(
    df['Итого сумма'].astype(str).str.replace(" ", "").str.replace(",", "."),
    errors='coerce'
).fillna(0)

# ===================== ВКЛАДКИ =====================
tab1, tab2, tab3 = st.tabs(["📊 Биллинг", "🚚 Наём", "⚖️ Сравнение"])

# ===================== БИЛЛИНГ =====================
with tab1:

    col1, col2 = st.columns(2)
    start = col1.date_input("Дата от", df['Дата'].min())
    end = col2.date_input("Дата до", df['Дата'].max())

    df_f = df[(df['Дата'] >= pd.to_datetime(start)) &
              (df['Дата'] <= pd.to_datetime(end))]

    # KPI
    income = df_f['Итого сумма'].sum()
    qty = df_f['Кол-во шт'].sum()

    c1, c2, c3 = st.columns(3)
    c1.metric("Доход", f"{int(income):,}")
    c2.metric("Шт", f"{int(qty):,}")
    c3.metric("Цена за шт", f"{int(income/qty) if qty else 0:,}")

    # ===================== ГРУППИРОВКА =====================
    grouped = df_f.groupby(['Дата','Тип ТС','Регион']).agg({
        'Кол-во шт': 'sum',
        'Итого сумма': 'sum',
        'Магазин': 'nunique'
    }).reset_index()

    grouped = grouped.rename(columns={'Магазин': 'Кол-во точек'})

    # 💥 БЕЗ apply (ВАЖНО)
    grouped['1 точка'] = (grouped['Кол-во точек'] == 1).astype(int)
    grouped['2+ точек'] = (grouped['Кол-во точек'] > 1).astype(int)

    # грузчик (пример)
    grouped['Грузчик'] = (grouped['Кол-во шт'] > 1000).astype(int)

    st.dataframe(grouped, use_container_width=True)

    # график %
    fig = px.pie(grouped, names='Регион', values='Кол-во шт')
    st.plotly_chart(fig, use_container_width=True)

# ===================== НАЁМ =====================
with tab2:

    url_hire = "https://docs.google.com/spreadsheets/d/1-kv25KvN60XPksMm4YC27r_x9cxn0PevY4npWWdOi4c/export?format=csv&gid=0"
    hire = pd.read_csv(url_hire)
    hire.columns = hire.columns.str.strip()

    if 'Итого сумма' in hire.columns:
        hire['Итого сумма'] = pd.to_numeric(
            hire['Итого сумма'].astype(str).str.replace(" ", ""),
            errors='coerce'
        ).fillna(0)

        st.metric("Расход", f"{int(hire['Итого сумма'].sum()):,}")

    st.dataframe(hire, use_container_width=True)

# ===================== СРАВНЕНИЕ =====================
with tab3:

    hire_total = hire['Итого сумма'].sum() if 'Итого сумма' in hire.columns else 0
    income = df['Итого сумма'].sum()

    profit = income - hire_total
    margin = (profit / income * 100) if income else 0

    c1, c2, c3 = st.columns(3)
    c1.metric("Доход", f"{int(income):,}")
    c2.metric("Наём", f"{int(hire_total):,}")
    c3.metric("Маржа %", f"{margin:.1f}%")

    chart = pd.DataFrame({
        'Тип': ['Доход', 'Расход'],
        'Сумма': [income, hire_total]
    })

    fig = px.bar(chart, x='Тип', y='Сумма')
    st.plotly_chart(fig, use_container_width=True)

# ===================== ВЫГРУЗКА =====================
st.subheader("📥 Выгрузка")

if st.button("Скачать Excel"):
    df.to_excel("report.xlsx", index=False)
    with open("report.xlsx", "rb") as f:
        st.download_button("Скачать файл", f, "report.xlsx")
