import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(layout="wide")
st.title("📊 Биллинг Магнит")

# ===================== ЗАГРУЗКА =====================
url = "https://docs.google.com/spreadsheets/d/1-kv25KvN60XPksMm4YC27r_x9cxn0PevY4npWWdOi4c/export?format=csv&gid=348938437"
df = pd.read_csv(url)

df.columns = df.columns.str.strip()

# ===================== ПОИСК КОЛОНОК =====================
def find_col(name):
    for col in df.columns:
        if name.lower() in col.lower():
            return col
    return None

date_col = find_col('дата')
qty_col = find_col('шт')
sum_col = find_col('сумм')
region_col = find_col('регион')
auto_col = find_col('авто')
shop_col = find_col('магаз')

# ===================== ПРОВЕРКА =====================
if not all([date_col, qty_col, sum_col]):
    st.error("❌ Не найдены нужные колонки")
    st.write(df.columns)
    st.stop()

# ===================== ПЕРЕИМЕНОВАНИЕ =====================
df = df.rename(columns={
    date_col: 'Дата',
    qty_col: 'Кол-во шт',
    sum_col: 'Итого сумма',
    region_col: 'Регион',
    auto_col: 'Тип ТС',
    shop_col: 'Магазин'
})

# ===================== ЧИСТКА ДАННЫХ =====================

# дата
df['Дата'] = pd.to_datetime(df['Дата'], errors='coerce')

# количество
df['Кол-во шт'] = (
    df['Кол-во шт']
    .astype(str)
    .str.replace(" ", "")
    .str.replace(",", ".")
)

df['Кол-во шт'] = pd.to_numeric(df['Кол-во шт'], errors='coerce').fillna(0)

# сумма
df['Итого сумма'] = (
    df['Итого сумма']
    .astype(str)
    .str.replace(" ", "")
    .str.replace(",", ".")
)

df['Итого сумма'] = pd.to_numeric(df['Итого сумма'], errors='coerce').fillna(0)

# ===================== ФИЛЬТР =====================
col1, col2 = st.columns(2)

start_date = col1.date_input("Дата от", df['Дата'].min())
end_date = col2.date_input("Дата до", df['Дата'].max())

df = df[(df['Дата'] >= pd.to_datetime(start_date)) &
        (df['Дата'] <= pd.to_datetime(end_date))]

# ===================== KPI =====================
total_income = float(df['Итого сумма'].sum())
total_qty = float(df['Кол-во шт'].sum())

price_per_unit = total_income / total_qty if total_qty else 0

c1, c2, c3 = st.columns(3)

c1.metric("💰 Доход", f"{int(total_income):,}")
c2.metric("📦 Шт", f"{int(total_qty):,}")
c3.metric("💸 Цена за шт", f"{int(price_per_unit):,}")

# ===================== ГРУППИРОВКА =====================
grouped = df.groupby(['Регион', 'Тип ТС']).agg({
    'Кол-во шт': 'sum',
    'Итого сумма': 'sum',
    'Магазин': 'nunique'
}).reset_index()

grouped = grouped.rename(columns={'Магазин': 'Кол-во точек'})

# 1 / 2+
grouped['1 точка'] = grouped['Кол-во точек'].apply(lambda x: 1 if x == 1 else 0)
grouped['2+ точек'] = grouped['Кол-во точек'].apply(lambda x: 1 if x > 1 else 0)

# грузчик (пример)
grouped['Грузчик'] = grouped['Кол-во шт'].apply(lambda x: 1 if x > 1000 else 0)

st.subheader("📋 Биллинг")
st.dataframe(grouped, use_container_width=True)

# ===================== ГРАФИК =====================
fig = px.pie(grouped, names='Регион', values='Кол-во шт')
st.plotly_chart(fig, use_container_width=True)
