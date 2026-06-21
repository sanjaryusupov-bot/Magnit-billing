import streamlit as st
import pandas as pd
import plotly.express as px

# ===================== НАСТРОЙКА =====================
st.set_page_config(layout="wide", page_title="Биллинг Магнит")

# 🎨 СТИЛЬ
st.markdown("""
<style>
body {
    background: linear-gradient(135deg, #0f2027, #203a43, #2c5364);
    color: white;
}
.metric {
    background: #1e2a38;
    padding: 15px;
    border-radius: 10px;
}
</style>
""", unsafe_allow_html=True)

st.title("📊 Биллинг с Магнитом")

# ===================== ЗАГРУЗКА =====================

url = "https://docs.google.com/spreadsheets/d/1-kv25KvN60XPksMm4YC27r_x9cxn0PevY4npWWdOi4c/export?format=csv&gid=348938437"

df = pd.read_csv(url)

# ===================== ПОДГОТОВКА =====================

df['Дата'] = pd.to_datetime(df['Дата'], errors='coerce')

# фильтр
start_date = st.date_input("Дата от", df['Дата'].min())
end_date = st.date_input("Дата до", df['Дата'].max())

df = df[(df['Дата'] >= pd.to_datetime(start_date)) &
        (df['Дата'] <= pd.to_datetime(end_date))]

# ===================== KPI =====================

total_income = df['Итого сумма'].sum()
total_qty = df['Кол-во шт'].sum()

price_per_unit = total_income / total_qty if total_qty > 0 else 0

col1, col2, col3 = st.columns(3)

col1.metric("💰 Доход", f"{total_income:,.0f}")
col2.metric("📦 Кол-во шт", f"{total_qty:,.0f}")
col3.metric("💸 Цена за шт", f"{price_per_unit:,.0f}")

# ===================== ГРУППИРОВКА =====================

group = df.groupby(['Регион', 'Тип ТС']).agg({
    'Кол-во шт': 'sum',
    'Итого сумма': 'sum'
}).reset_index()

group['Цена за шт'] = group['Итого сумма'] / group['Кол-во шт']

# ===================== ГРАФИКИ =====================

st.subheader("📍 Распределение по регионам")

fig1 = px.pie(group, names='Регион', values='Кол-во шт')
st.plotly_chart(fig1, use_container_width=True)

st.subheader("🚚 По типу транспорта")

fig2 = px.bar(group, x='Тип ТС', y='Кол-во шт', color='Регион')
st.plotly_chart(fig2, use_container_width=True)

# ===================== СРАВНЕНИЕ С НАЁМОМ =====================

st.subheader("⚖️ Сравнение с наёмом")

# пример (замени на свою вкладку)
hire_url = "https://docs.google.com/spreadsheets/d/1-kv25KvN60XPksMm4YC27r_x9cxn0PevY4npWWdOi4c/export?format=csv&gid=0"

hire = pd.read_csv(hire_url)

hire_total = hire['Итого сумма'].sum()

profit = total_income - hire_total
margin = (profit / total_income * 100) if total_income else 0

c1, c2, c3 = st.columns(3)

c1.metric("Доход (Магнит)", f"{total_income:,.0f}")
c2.metric("Расход (наём)", f"{hire_total:,.0f}")
c3.metric("Маржа %", f"{margin:.1f}%")

# ===================== ГРАФИК МАРЖИ =====================

compare_df = pd.DataFrame({
    'Тип': ['Доход', 'Расход'],
    'Сумма': [total_income, hire_total]
})

fig3 = px.bar(compare_df, x='Тип', y='Сумма')
st.plotly_chart(fig3, use_container_width=True)

# ===================== ВЫГРУЗКА =====================

st.subheader("📥 Выгрузка")

if st.button("Скачать Excel"):
    df.to_excel("report.xlsx", index=False)
    with open("report.xlsx", "rb") as f:
        st.download_button("Скачать файл", f, "report.xlsx")
