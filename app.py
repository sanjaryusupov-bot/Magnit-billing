import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(layout="wide", page_title="Биллинг Магнит")

st.title("📊 Биллинг с Магнитом")

# ===================== ЗАГРУЗКА =====================

url = "https://docs.google.com/spreadsheets/d/1-kv25KvN60XPksMm4YC27r_x9cxn0PevY4npWWdOi4c/export?format=csv&gid=348938437"
df = pd.read_csv(url)

# чистка колонок
df.columns = df.columns.str.strip()

# 👇 ПОКАЗЫВАЕМ ЧТО ПРИШЛО (разово)
st.write("Колонки:", df.columns)

# ===================== АВТООПРЕДЕЛЕНИЕ КОЛОНОК =====================

def find_col(possible_names):
    for name in possible_names:
        for col in df.columns:
            if name.lower() in col.lower():
                return col
    return None

date_col = find_col(['дата'])
qty_col = find_col(['шт'])
sum_col = find_col(['сумм'])
region_col = find_col(['регион'])
auto_col = find_col(['авто'])
shop_col = find_col(['магаз'])

# проверка
if not date_col:
    st.error("❌ Не найдена колонка с датой")
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

# дата
df['Дата'] = pd.to_datetime(df['Дата'], errors='coerce')

# ===================== ВКЛАДКИ =====================

tab1, tab2, tab3 = st.tabs(["📊 Общий", "🚚 Наём", "⚖️ Сравнение"])

# ===================== ОБЩИЙ =====================

with tab1:

    st.subheader("Фильтры")

    col1, col2 = st.columns(2)
    start_date = col1.date_input("Дата от", df['Дата'].min())
    end_date = col2.date_input("Дата до", df['Дата'].max())

    if 'Подрядчик' in df.columns:
        contractor = st.selectbox("Подрядчик", ["Все"] + list(df['Подрядчик'].dropna().unique()))
        if contractor != "Все":
            df = df[df['Подрядчик'] == contractor]

    df_f = df[(df['Дата'] >= pd.to_datetime(start_date)) &
              (df['Дата'] <= pd.to_datetime(end_date))]

    # KPI
    total_income = df_f['Итого сумма'].sum()
    total_qty = df_f['Кол-во шт'].sum()

    c1, c2, c3 = st.columns(3)
    c1.metric("💰 Доход", f"{total_income:,.0f}")
    c2.metric("📦 Шт", f"{total_qty:,.0f}")
    c3.metric("💸 Цена за шт", f"{(total_income/total_qty if total_qty else 0):,.0f}")

    # ===================== 1 ТОЧКА / 2+ =====================

    grouped = df_f.groupby(['Дата', 'Тип ТС', 'Регион']).agg({
        'Магазин': 'nunique',
        'Кол-во шт': 'sum',
        'Итого сумма': 'sum'
    }).reset_index()

    grouped = grouped.rename(columns={'Магазин': 'Кол-во точек'})

    grouped['1 точка'] = grouped['Кол-во точек'].apply(lambda x: 1 if x == 1 else 0)
    grouped['2+ точек'] = grouped['Кол-во точек'].apply(lambda x: 1 if x > 1 else 0)

    # ГРУЗЧИК (пример логики — можно поменять)
    grouped['Грузчик'] = grouped['Кол-во шт'].apply(lambda x: 1 if x > 1000 else 0)

    st.subheader("📋 Расчет биллинга")
    st.dataframe(grouped, use_container_width=True)

    # график
    fig = px.pie(grouped, names='Регион', values='Кол-во шт')
    st.plotly_chart(fig, use_container_width=True)


# ===================== НАЁМ =====================

with tab2:

    st.subheader("🚚 Наём")

    hire_url = "https://docs.google.com/spreadsheets/d/1-kv25KvN60XPksMm4YC27r_x9cxn0PevY4npWWdOi4c/export?format=csv&gid=0"
    hire = pd.read_csv(hire_url)
    hire.columns = hire.columns.str.strip()

    st.write("Колонки (наём):", hire.columns)

    if 'Итого сумма' in hire.columns:
        st.metric("Расход на наём", f"{hire['Итого сумма'].sum():,.0f}")
    else:
        st.warning("Нет колонки 'Итого сумма'")

    st.dataframe(hire, use_container_width=True)


# ===================== СРАВНЕНИЕ =====================

with tab3:

    st.subheader("⚖️ Сравнение")

    hire_total = hire['Итого сумма'].sum() if 'Итого сумма' in hire.columns else 0
    income = df['Итого сумма'].sum()

    profit = income - hire_total
    margin = (profit / income * 100) if income else 0

    c1, c2, c3 = st.columns(3)
    c1.metric("Доход", f"{income:,.0f}")
    c2.metric("Наём", f"{hire_total:,.0f}")
    c3.metric("Маржа %", f"{margin:.1f}%")

    chart = pd.DataFrame({
        'Тип': ['Доход', 'Расход'],
        'Сумма': [income, hire_total]
    })

    fig = px.bar(chart, x='Тип', y='Сумма')
    st.plotly_chart(fig, use_container_width=True)
