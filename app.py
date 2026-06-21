import streamlit as st
import pandas as pd
import plotly.express as px

# ===================== НАСТРОЙКА =====================
st.set_page_config(layout="wide", page_title="Биллинг Магнит")

st.markdown("""
<style>
body {
    background: linear-gradient(135deg, #0f2027, #203a43, #2c5364);
    color: white;
}
.block {
    background: #1e2a38;
    padding: 15px;
    border-radius: 10px;
}
</style>
""", unsafe_allow_html=True)

st.title("📊 Биллинг с Магнитом")

# ===================== ЗАГРУЗКА ДАННЫХ =====================

# Общая таблица (клиент)
url_main = "https://docs.google.com/spreadsheets/d/1-kv25KvN60XPksMm4YC27r_x9cxn0PevY4npWWdOi4c/export?format=csv&gid=348938437"
df = pd.read_csv(url_main)

# чистим названия
df.columns = df.columns.str.strip()

# ПЕРЕИМЕНОВАНИЕ (ВАЖНО)
df = df.rename(columns={
    'Дата отгрузки': 'Дата',
    'Кол-во шт': 'Кол-во шт',
    'Итого сумма': 'Итого сумма',
    'Регион': 'Регион',
    'Авто': 'Тип ТС',
    'Магазин': 'Магазин'
})

# проверка
required_cols = ['Дата', 'Кол-во шт', 'Итого сумма']
for col in required_cols:
    if col not in df.columns:
        st.error(f"Нет колонки: {col}")
        st.write(df.columns)
        st.stop()

# дата
df['Дата'] = pd.to_datetime(df['Дата'], errors='coerce')

# ===================== ФИЛЬТР =====================

col1, col2 = st.columns(2)

start_date = col1.date_input("Дата от", df['Дата'].min())
end_date = col2.date_input("Дата до", df['Дата'].max())

df = df[(df['Дата'] >= pd.to_datetime(start_date)) &
        (df['Дата'] <= pd.to_datetime(end_date))]

# ===================== KPI =====================

total_income = df['Итого сумма'].sum()
total_qty = df['Кол-во шт'].sum()
price_per_unit = total_income / total_qty if total_qty > 0 else 0

c1, c2, c3 = st.columns(3)

c1.metric("💰 Доход (Магнит)", f"{total_income:,.0f}")
c2.metric("📦 Кол-во шт", f"{total_qty:,.0f}")
c3.metric("💸 Цена за шт", f"{price_per_unit:,.0f}")

# ===================== ГРУППИРОВКА =====================

group = df.groupby(['Регион', 'Тип ТС']).agg({
    'Кол-во шт': 'sum',
    'Итого сумма': 'sum',
    'Магазин': 'nunique'
}).reset_index()

group = group.rename(columns={'Магазин': 'Кол-во точек'})

group['Цена за шт'] = group['Итого сумма'] / group['Кол-во шт']

st.subheader("📊 Таблица аналитики")
st.dataframe(group, use_container_width=True)

# ===================== ГРАФИКИ =====================

st.subheader("📍 Распределение по регионам")
fig1 = px.pie(group, names='Регион', values='Кол-во шт')
st.plotly_chart(fig1, use_container_width=True)

st.subheader("🚚 По типу транспорта")
fig2 = px.bar(group, x='Тип ТС', y='Кол-во шт', color='Регион')
st.plotly_chart(fig2, use_container_width=True)

# ===================== НАЁМ =====================

st.subheader("⚖️ Сравнение с наёмом")

url_hire = "https://docs.google.com/spreadsheets/d/1-kv25KvN60XPksMm4YC27r_x9cxn0PevY4npWWdOi4c/export?format=csv&gid=0"
hire = pd.read_csv(url_hire)

hire.columns = hire.columns.str.strip()

# ожидаемые колонки
if 'Итого сумма' not in hire.columns:
    st.warning("Проверь вкладку 'Сводная по дням' (нет Итого сумма)")
    hire_total = 0
else:
    hire_total = hire['Итого сумма'].sum()

profit = total_income - hire_total
margin = (profit / total_income * 100) if total_income else 0

h1, h2, h3 = st.columns(3)

h1.metric("Доход", f"{total_income:,.0f}")
h2.metric("Расход (наём)", f"{hire_total:,.0f}")
h3.metric("Маржа %", f"{margin:.1f}%")

# ===================== ГРАФИК СРАВНЕНИЯ =====================

compare_df = pd.DataFrame({
    'Тип': ['Доход', 'Расход'],
    'Сумма': [total_income, hire_total]
})

fig3 = px.bar(compare_df, x='Тип', y='Сумма')
st.plotly_chart(fig3, use_container_width=True)

# ===================== ВЫГРУЗКА =====================

st.subheader("📥 Выгрузка")

@st.cache_data
def convert_to_excel(dataframe):
    return dataframe.to_excel(index=False)

if st.button("Скачать отчет"):
    file = df.to_excel("report.xlsx", index=False)
    with open("report.xlsx", "rb") as f:
        st.download_button("Скачать Excel", f, file_name="report.xlsx")
