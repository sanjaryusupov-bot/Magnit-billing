import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import io
import base64
import json
import numpy as np

# Попытка импорта Google Sheets API с обработкой ошибок
try:
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    GOOGLE_SHEETS_AVAILABLE = True
except ImportError:
    GOOGLE_SHEETS_AVAILABLE = False
    st.warning("Библиотеки Google Sheets не установлены. Используются тестовые данные.")

# Настройка страницы
st.set_page_config(
    page_title="Аналитический дашборд логистики",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Стили для красивого фона
st.markdown("""
    <style>
        .stApp {
            background: linear-gradient(135deg, #0f0c29, #302b63, #24243e);
            color: white;
        }
        .css-1d391kg {
            background: rgba(255, 255, 255, 0.05);
            backdrop-filter: blur(10px);
            border-radius: 15px;
            border: 1px solid rgba(255, 255, 255, 0.1);
        }
        .metric-card {
            background: rgba(255, 255, 255, 0.05);
            backdrop-filter: blur(10px);
            border-radius: 15px;
            padding: 20px;
            border: 1px solid rgba(255, 255, 255, 0.1);
            box-shadow: 0 8px 32px rgba(0,0,0,0.3);
            transition: transform 0.3s ease;
        }
        .metric-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 12px 48px rgba(0,0,0,0.4);
        }
        .metric-value {
            font-size: 2.5rem;
            font-weight: bold;
            color: #00f5ff;
            text-shadow: 0 0 20px rgba(0, 245, 255, 0.3);
        }
        .metric-label {
            color: rgba(255,255,255,0.7);
            font-size: 0.9rem;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        .glass-card {
            background: rgba(255, 255, 255, 0.05);
            backdrop-filter: blur(10px);
            border-radius: 15px;
            padding: 20px;
            border: 1px solid rgba(255, 255, 255, 0.1);
            box-shadow: 0 8px 32px rgba(0,0,0,0.2);
        }
        .st-emotion-cache-1r6slb0 {
            background: rgba(255, 255, 255, 0.05);
            backdrop-filter: blur(10px);
        }
        .st-emotion-cache-16idsys {
            background: rgba(255, 255, 255, 0.03);
        }
        h1, h2, h3 {
            color: #00f5ff !important;
            text-shadow: 0 0 20px rgba(0, 245, 255, 0.2);
        }
        .st-emotion-cache-1y4p8pa {
            max-width: 100% !important;
        }
        .stButton > button {
            background: linear-gradient(135deg, #00f5ff, #00b4ff);
            color: white;
            border: none;
            border-radius: 10px;
            padding: 10px 25px;
            font-weight: bold;
            transition: all 0.3s ease;
            box-shadow: 0 4px 15px rgba(0, 245, 255, 0.3);
        }
        .stButton > button:hover {
            transform: scale(1.05);
            box-shadow: 0 8px 25px rgba(0, 245, 255, 0.5);
        }
        .st-emotion-cache-183lzff {
            background: rgba(255, 255, 255, 0.05);
        }
        .st-emotion-cache-1v0mbdj {
            color: white;
        }
        /* Стиль для кастомных индикаторов */
        .indicator-label {
            color: rgba(255,255,255,0.7);
            font-size: 0.85rem;
        }
        .indicator-value {
            font-size: 1.8rem;
            font-weight: bold;
        }
        .indicator-delta {
            font-size: 1rem;
            font-weight: bold;
        }
        .indicator-delta-positive {
            color: #00ff88;
        }
        .indicator-delta-negative {
            color: #ff6b6b;
        }
    </style>
""", unsafe_allow_html=True)

# Функция для генерации тестовых данных
def generate_test_data():
    np.random.seed(42)
    dates = pd.date_range('2026-01-01', '2026-06-21', freq='D')
    
    data = []
    regions = ['Центральный', 'Северо-Западный', 'Южный', 'Приволжский', 'Уральский']
    transport_types = ['Грузовик', 'Фургон', 'Рефрижератор']
    stores = ['Магнит А', 'Магнит Б', 'Магнит В', 'Магнит Г', 'Магнит Д']
    
    for date in dates:
        # Генерируем от 1 до 5 записей в день
        for _ in range(np.random.randint(1, 5)):
            region = np.random.choice(regions)
            transport = np.random.choice(transport_types)
            store = np.random.choice(stores)
            quantity = np.random.randint(100, 500)
            # Разная цена для биллинга и найма
            if np.random.random() > 0.5:
                price_per_unit = np.random.uniform(80, 150)
                type_ = 'Биллинг'
            else:
                price_per_unit = np.random.uniform(50, 100)
                type_ = 'Наём'
            
            amount = quantity * price_per_unit
            
            data.append({
                'Дата': date,
                'Тип ТС': transport,
                'Регион': region,
                'Магазин': store,
                'Кол-во шт': quantity,
                'Сумма': round(amount, 2),
                'Тип': type_
            })
    
    return pd.DataFrame(data)

# Загрузка данных из Google Sheets или генерация тестовых
@st.cache_data(ttl=3600)
def load_data():
    try:
        if GOOGLE_SHEETS_AVAILABLE:
            # Настройка доступа к Google Sheets
            scope = ['https://spreadsheets.google.com/feeds',
                    'https://www.googleapis.com/auth/drive']
            
            # Пытаемся загрузить из secrets
            try:
                creds_dict = {
                    "type": st.secrets["google"]["type"],
                    "project_id": st.secrets["google"]["project_id"],
                    "private_key_id": st.secrets["google"]["private_key_id"],
                    "private_key": st.secrets["google"]["private_key"],
                    "client_email": st.secrets["google"]["client_email"],
                    "client_id": st.secrets["google"]["client_id"],
                    "auth_uri": st.secrets["google"]["auth_uri"],
                    "token_uri": st.secrets["google"]["token_uri"],
                    "auth_provider_x509_cert_url": st.secrets["google"]["auth_provider_x509_cert_url"],
                    "client_x509_cert_url": st.secrets["google"]["client_x509_cert_url"]
                }
                creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
                client = gspread.authorize(creds)
                sheet = client.open_by_key('1-kv25KvN60XPksMm4YC27r_x9cxn0PevY4npWWdOi4c')
                
                # Загружаем данные из листа "Общ таблица"
                worksheet = sheet.worksheet('Общ таблица')
                data = worksheet.get_all_values()
                
                if len(data) > 1:
                    headers = data[0]
                    rows = data[1:]
                    df = pd.DataFrame(rows, columns=headers)
                    
                    # Преобразуем типы данных
                    df['Дата'] = pd.to_datetime(df['Дата'])
                    df['Кол-во шт'] = pd.to_numeric(df['Кол-во шт'], errors='coerce')
                    df['Сумма'] = pd.to_numeric(df['Сумма'], errors='coerce')
                    
                    # Добавляем тип (Биллинг или Наём)
                    df['Тип'] = 'Биллинг'  # По умолчанию
                    
                    return df
                else:
                    st.warning("Таблица пуста. Используются тестовые данные.")
                    return generate_test_data()
                    
            except Exception as e:
                st.warning(f"Ошибка доступа к Google Sheets: {e}. Используются тестовые данные.")
                return generate_test_data()
        else:
            st.info("Библиотеки Google Sheets не установлены. Используются демонстрационные данные.")
            return generate_test_data()
            
    except Exception as e:
        st.error(f"Ошибка загрузки данных: {e}")
        return generate_test_data()

# Функция для создания выгрузки в Excel
def create_excel_download(df, filename="analytics_data.xlsx"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Данные', index=False)
        
        # Добавляем сводные таблицы
        if 'Тип ТС' in df.columns:
            pivot = df.pivot_table(
                values='Кол-во шт',
                index=['Дата', 'Регион'],
                columns='Тип ТС',
                aggfunc='sum',
                fill_value=0
            )
            pivot.to_excel(writer, sheet_name='Сводка по ТС')
        
        if 'Тип' in df.columns:
            pivot_type = df.pivot_table(
                values='Сумма',
                index=['Дата'],
                columns='Тип',
                aggfunc='sum',
                fill_value=0
            )
            pivot_type.to_excel(writer, sheet_name='Сводка по типу')
        
        pivot_region = df.pivot_table(
            values='Сумма',
            index=['Дата'],
            columns='Регион',
            aggfunc='sum',
            fill_value=0
        )
        pivot_region.to_excel(writer, sheet_name='Сводка по регионам')
    
    output.seek(0)
    return output

# Основная логика дашборда
def main():
    # Заголовок
    st.title("🚛 Аналитический дашборд логистики")
    st.markdown("*Сравнительный анализ Биллинг vs Наём*")
    
    # Загрузка данных
    with st.spinner("Загрузка данных..."):
        df = load_data()
    
    if df.empty:
        st.warning("Нет данных для отображения")
        return
    
    # Преобразование дат
    df['Дата'] = pd.to_datetime(df['Дата'])
    
    # Боковая панель с фильтрами
    st.sidebar.markdown("## 🎯 Фильтры")
    st.sidebar.markdown("---")
    
    # Выбор периода
    min_date = df['Дата'].min().date()
    max_date = df['Дата'].max().date()
    
    date_range = st.sidebar.date_input(
        "📅 Период",
        [min_date, max_date],
        min_value=min_date,
        max_value=max_date
    )
    
    if len(date_range) == 2:
        start_date, end_date = date_range
        mask = (df['Дата'].dt.date >= start_date) & (df['Дата'].dt.date <= end_date)
        filtered_df = df.loc[mask]
    else:
        filtered_df = df
    
    # Фильтры
    if 'Регион' in filtered_df.columns:
        regions = st.sidebar.multiselect(
            "📍 Регионы",
            options=sorted(filtered_df['Регион'].unique()),
            default=sorted(filtered_df['Регион'].unique())
        )
        
        if regions:
            filtered_df = filtered_df[filtered_df['Регион'].isin(regions)]
    
    if 'Тип ТС' in filtered_df.columns:
        transport = st.sidebar.multiselect(
            "🚚 Тип ТС",
            options=sorted(filtered_df['Тип ТС'].unique()),
            default=sorted(filtered_df['Тип ТС'].unique())
        )
        
        if transport:
            filtered_df = filtered_df[filtered_df['Тип ТС'].isin(transport)]
    
    # Кнопка обновления
    if st.sidebar.button("🔄 Обновить данные", type="primary"):
        st.cache_data.clear()
        st.rerun()
    
    st.sidebar.markdown("---")
    st.sidebar.markdown("**ℹ️ Информация**")
    st.sidebar.markdown(f"Всего записей: {len(filtered_df):,}")
    st.sidebar.markdown(f"Период: {filtered_df['Дата'].min().date()} - {filtered_df['Дата'].max().date()}")
    
    # Основные метрики
    st.markdown("## 📊 Ключевые метрики")
    
    col1, col2, col3, col4 = st.columns(4)
    
    # Общие метрики
    total_quantity = filtered_df['Кол-во шт'].sum()
    total_amount = filtered_df['Сумма'].sum()
    avg_amount_per_unit = total_amount / total_quantity if total_quantity > 0 else 0
    
    # Метрики по типу
    if 'Тип' in filtered_df.columns:
        billing_df = filtered_df[filtered_df['Тип'] == 'Биллинг']
        hired_df = filtered_df[filtered_df['Тип'] == 'Наём']
    else:
        # Если нет типа, считаем все как биллинг
        billing_df = filtered_df
        hired_df = pd.DataFrame()
    
    billing_quantity = billing_df['Кол-во шт'].sum() if not billing_df.empty else 0
    hired_quantity = hired_df['Кол-во шт'].sum() if not hired_df.empty else 0
    billing_amount = billing_df['Сумма'].sum() if not billing_df.empty else 0
    hired_amount = hired_df['Сумма'].sum() if not hired_df.empty else 0
    
    with col1:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">📦 Всего отгружено</div>
                <div class="metric-value">{total_quantity:,.0f} шт</div>
                <div style="color: rgba(255,255,255,0.5); font-size: 0.8rem;">на сумму {total_amount:,.0f} ₽</div>
                <div style="color: rgba(255,255,255,0.5); font-size: 0.8rem;">средняя цена за шт: {avg_amount_per_unit:,.2f} ₽</div>
            </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">💰 Биллинг (клиент)</div>
                <div class="metric-value">{billing_quantity:,.0f} шт</div>
                <div style="color: #00f5ff; font-size: 0.9rem;">{billing_amount:,.0f} ₽</div>
                <div style="color: rgba(255,255,255,0.5); font-size: 0.8rem;">средняя цена за шт: {billing_amount/billing_quantity if billing_quantity > 0 else 0:,.2f} ₽</div>
            </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">👷 Наём (исполнители)</div>
                <div class="metric-value">{hired_quantity:,.0f} шт</div>
                <div style="color: #ff6b6b; font-size: 0.9rem;">{hired_amount:,.0f} ₽</div>
                <div style="color: rgba(255,255,255,0.5); font-size: 0.8rem;">средняя цена за шт: {hired_amount/hired_quantity if hired_quantity > 0 else 0:,.2f} ₽</div>
            </div>
        """, unsafe_allow_html=True)
    
    with col4:
        profit = billing_amount - hired_amount
        profit_per_unit = profit / total_quantity if total_quantity > 0 else 0
        margin_percent = (profit / billing_amount * 100) if billing_amount > 0 else 0
        
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">📈 Маржинальная прибыль</div>
                <div class="metric-value" style="color: {'#00ff88' if profit >= 0 else '#ff6b6b'}">
                    {profit:,.0f} ₽
                </div>
                <div style="color: rgba(255,255,255,0.5); font-size: 0.8rem;">
                    прибыль за шт: {profit_per_unit:,.2f} ₽
                </div>
                <div style="color: rgba(255,255,255,0.5); font-size: 0.8rem;">
                    маржинальность: {margin_percent:.1f}%
                </div>
            </div>
        """, unsafe_allow_html=True)
    
    # Графики в 2 колонки
    st.markdown("## 📈 Аналитические графики")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.subheader("Сравнение Биллинг vs Наём")
        
        comparison_data = pd.DataFrame({
            'Показатель': ['Кол-во шт', 'Сумма (млн ₽)'],
            'Биллинг': [billing_quantity, billing_amount / 1000000],
            'Наём': [hired_quantity, hired_amount / 1000000]
        })
        
        fig = go.Figure()
        fig.add_trace(go.Bar(
            name='Биллинг',
            x=comparison_data['Показатель'],
            y=comparison_data['Биллинг'],
            marker_color='#00f5ff',
            text=comparison_data['Биллинг'].round(1),
            textposition='outside',
            textfont=dict(color='white'),
            marker=dict(
                pattern_shape='solid',
                pattern_fillmode='replace'
            )
        ))
        fig.add_trace(go.Bar(
            name='Наём',
            x=comparison_data['Показатель'],
            y=comparison_data['Наём'],
            marker_color='#ff6b6b',
            text=comparison_data['Наём'].round(1),
            textposition='outside',
            textfont=dict(color='white')
        ))
        
        fig.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white',
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            ),
            margin=dict(l=0, r=0, t=30, b=0),
            height=350,
            xaxis=dict(gridcolor='rgba(255,255,255,0.1)'),
            yaxis=dict(gridcolor='rgba(255,255,255,0.1)'),
            bargap=0.3,
            bargroupgap=0.1
        )
        st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.subheader("Распределение по регионам")
        
        if 'Регион' in filtered_df.columns:
            region_stats = filtered_df.groupby('Регион').agg({
                'Кол-во шт': 'sum',
                'Сумма': 'sum'
            }).reset_index()
            region_stats['Средняя цена'] = region_stats['Сумма'] / region_stats['Кол-во шт']
            
            fig = px.pie(
                region_stats,
                values='Кол-во шт',
                names='Регион',
                title='Доля по регионам',
                color_discrete_sequence=px.colors.sequential.RdBu[::-1],
                hole=0.4
            )
            fig.update_traces(
                textposition='inside',
                textinfo='percent+label',
                textfont=dict(color='white', size=12)
            )
            fig.update_layout(
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font_color='white',
                showlegend=False,
                margin=dict(l=0, r=0, t=30, b=0),
                height=350
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Нет данных по регионам")
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Динамика по дням
    st.markdown("## 📉 Динамика показателей")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.subheader("Динамика отгрузок")
        
        daily_stats = filtered_df.groupby('Дата').agg({
            'Кол-во шт': 'sum',
            'Сумма': 'sum'
        }).reset_index()
        
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=daily_stats['Дата'],
            y=daily_stats['Кол-во шт'],
            mode='lines+markers',
            name='Кол-во шт',
            line=dict(color='#00f5ff', width=3),
            marker=dict(size=6, color='#00f5ff', symbol='circle')
        ))
        fig.add_trace(go.Scatter(
            x=daily_stats['Дата'],
            y=daily_stats['Сумма'],
            mode='lines+markers',
            name='Сумма ₽',
            yaxis='y2',
            line=dict(color='#ffd93d', width=3),
            marker=dict(size=6, color='#ffd93d', symbol='diamond')
        ))
        
        fig.update_layout(
            yaxis=dict(
                title="Кол-во шт",
                titlefont=dict(color='#00f5ff'),
                tickfont=dict(color='#00f5ff'),
                gridcolor='rgba(255,255,255,0.1)'
            ),
            yaxis2=dict(
                title="Сумма ₽",
                titlefont=dict(color='#ffd93d'),
                tickfont=dict(color='#ffd93d'),
                overlaying='y',
                side='right',
                gridcolor='rgba(255,255,255,0.1)'
            ),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white',
            hovermode='x unified',
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            ),
            margin=dict(l=0, r=0, t=30, b=0),
            height=350,
            xaxis=dict(gridcolor='rgba(255,255,255,0.1)')
        )
        st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.subheader("Распределение по типам ТС")
        
        if 'Тип ТС' in filtered_df.columns:
            transport_stats = filtered_df.groupby('Тип ТС').agg({
                'Кол-во шт': 'sum',
                'Сумма': 'sum'
            }).reset_index()
            
            fig = px.bar(
                transport_stats,
                x='Тип ТС',
                y='Кол-во шт',
                color='Тип ТС',
                text='Кол-во шт',
                color_discrete_sequence=px.colors.sequential.Plasma_r,
                title='Объем перевозок по типам ТС'
            )
            fig.update_traces(textposition='outside', textfont_color='white')
            fig.update_layout(
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font_color='white',
                showlegend=False,
                margin=dict(l=0, r=0, t=30, b=0),
                height=350,
                xaxis=dict(gridcolor='rgba(255,255,255,0.1)'),
                yaxis=dict(gridcolor='rgba(255,255,255,0.1)')
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Нет данных по типам ТС")
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Анализ эффективности
    st.markdown("## 💰 Анализ эффективности")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.subheader("Средняя цена за шт")
        
        billing_avg = billing_amount / billing_quantity if billing_quantity > 0 else 0
        hired_avg = hired_amount / hired_quantity if hired_quantity > 0 else 0
        
        if billing_avg > 0 or hired_avg > 0:
            max_val = max(billing_avg, hired_avg) * 1.3
            
            fig = go.Figure()
            fig.add_trace(go.Indicator(
                mode="gauge+number+delta",
                value=billing_avg,
                delta={'reference': hired_avg, 'increasing': {'color': "#00ff88"}, 'decreasing': {'color': "#ff6b6b"}},
                gauge={
                    'axis': {'range': [0, max_val], 'tickfont': {'color': 'white'}},
                    'bar': {'color': "#00f5ff"},
                    'steps': [
                        {'range': [0, hired_avg], 'color': 'rgba(255,107,107,0.3)'},
                        {'range': [hired_avg, max_val], 'color': 'rgba(0,245,255,0.3)'}
                    ],
                    'threshold': {
                        'line': {'color': "#ff6b6b", 'width': 4},
                        'thickness': 0.75,
                        'value': hired_avg
                    }
                },
                number={'suffix': " ₽", 'font': {'color': 'white', 'size': 40}},
                title={'text': "Биллинг vs Наём", 'font': {'color': 'white', 'size': 16}}
            ))
            fig.update_layout(
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font_color='white',
                height=300,
                margin=dict(l=20, r=20, t=50, b=20)
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Нет данных для сравнения")
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.subheader("Маржинальность")
        
        margin = ((billing_amount - hired_amount) / billing_amount * 100) if billing_amount > 0 else 0
        
        fig = go.Figure(go.Indicator(
            mode="gauge+number",
            value=margin,
            title={'text': "Маржинальность %", 'font': {'color': 'white', 'size': 16}},
            number={'suffix': "%", 'font': {'color': 'white', 'size': 40}},
            gauge={
                'axis': {'range': [-20, 80], 'tickfont': {'color': 'white'}},
                'bar': {'color': "#00ff88" if margin > 20 else "#ffd93d" if margin > 0 else "#ff6b6b"},
                'steps': [
                    {'range': [-20, 0], 'color': 'rgba(255,107,107,0.2)'},
                    {'range': [0, 30], 'color': 'rgba(255,217,61,0.2)'},
                    {'range': [30, 80], 'color': 'rgba(0,255,136,0.2)'}
                ],
                'threshold': {
                    'line': {'color': "white", 'width': 4},
                    'thickness': 0.75,
                    'value': 20
                }
            }
        ))
        fig.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white',
            height=300,
            margin=dict(l=20, r=20, t=50, b=20)
        )
        st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col3:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.subheader("Эффективность перевозок")
        
        # Расчет эффективности (шт на 1 рубль)
        billing_efficiency = billing_quantity / billing_amount if billing_amount > 0 else 0
        hired_efficiency = hired_quantity / hired_amount if hired_amount > 0 else 0
        
        efficiency_df = pd.DataFrame({
            'Тип': ['Биллинг', 'Наём'],
            'Эффективность (шт/₽)': [billing_efficiency * 100, hired_efficiency * 100]
        })
        
        fig = px.bar(
            efficiency_df,
            x='Тип',
            y='Эффективность (шт/₽)',
            color='Тип',
            text='Эффективность (шт/₽)',
            color_discrete_sequence=['#00f5ff', '#ff6b6b']
        )
        fig.update_traces(
            texttemplate='%{text:.3f}',
            textposition='outside',
            textfont_color='white'
        )
        fig.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white',
            showlegend=False,
            height=300,
            margin=dict(l=0, r=0, t=30, b=0),
            xaxis=dict(gridcolor='rgba(255,255,255,0.1)'),
            yaxis=dict(gridcolor='rgba(255,255,255,0.1)')
        )
        st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Таблица с данными
    st.markdown("## 📋 Детализация данных")
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        # Выбор периода для детализации
        detail_date = st.date_input(
            "📅 Выберите дату для детализации",
            value=filtered_df['Дата'].max().date() if not filtered_df.empty else datetime.now().date(),
            min_value=filtered_df['Дата'].min().date() if not filtered_df.empty else datetime.now().date() - timedelta(days=30),
            max_value=filtered_df['Дата'].max().date() if not filtered_df.empty else datetime.now().date()
        )
    
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("📥 Выгрузить в Excel", type="primary"):
            excel_file = create_excel_download(filtered_df)
            b64 = base64.b64encode(excel_file.getvalue()).decode()
            href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="analytics_data.xlsx">📊 Скачать Excel</a>'
            st.markdown(href, unsafe_allow_html=True)
    
    # Детальная таблица
    if not filtered_df.empty:
        detail_df = filtered_df[filtered_df['Дата'].dt.date == detail_date]
        
        if not detail_df.empty:
            # Группировка для детализации
            detail_pivot = detail_df.groupby(['Регион', 'Тип ТС', 'Тип']).agg({
                'Магазин': lambda x: ', '.join(x.unique()),
                'Кол-во шт': 'sum',
                'Сумма': 'sum'
            }).reset_index()
            
            st.dataframe(
                detail_pivot,
                use_container_width=True,
                column_config={
                    'Регион': '📍 Регион',
                    'Тип ТС': '🚚 Тип ТС',
                    'Тип': '📌 Тип',
                    'Магазин': '🏪 Магазины',
                    'Кол-во шт': st.column_config.NumberColumn('📦 Кол-во шт', format='%d'),
                    'Сумма': st.column_config.NumberColumn('💰 Сумма ₽', format='% .2f')
                },
                hide_index=True
            )
        else:
            st.info("📭 Нет данных за выбранную дату")
    else:
        st.info("📭 Нет данных для отображения")
    
    # Дополнительная информация
    with st.expander("📌 Дополнительная информация", expanded=False):
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
                **💡 Аналитические выводы:**
                - Сравнительный анализ показывает разницу между доходами от клиентов и расходами на наём
                - Графики демонстрируют распределение по регионам и типам транспортных средств
                - Маржинальность показывает эффективность бизнес-модели
                - Динамика позволяет отслеживать тренды и сезонность
            """)
        
        with col2:
            st.markdown("""
                **📊 Структура данных:**
                - **Биллинг** - доходы от клиента (Магнит)
                - **Наём** - расходы на исполнителей
                - **Маржинальная прибыль** = Биллинг - Наём
                - **Эффективность** = Кол-во шт / Сумма
            """)
    
    # Footer
    st.markdown("---")
    st.markdown(
        """
        <div style='text-align: center; color: rgba(255,255,255,0.3); font-size: 0.8rem;'>
            🚛 Аналитический дашборд логистики | Данные обновлены: {}
        </div>
        """.format(datetime.now().strftime("%d.%m.%Y %H:%M")),
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
