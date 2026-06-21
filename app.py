import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import io
import base64
import numpy as np

# Настройка страницы
st.set_page_config(
    page_title="Аналитика логистики | Биллинг vs Наём",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Стили для красивого фона
st.markdown("""
    <style>
        .stApp {
            background: linear-gradient(145deg, #0a0e27, #1a1f3a, #0d1230);
            color: white;
        }
        .main-header {
            text-align: center;
            padding: 20px;
            background: linear-gradient(135deg, rgba(0,245,255,0.1), rgba(0,180,255,0.05));
            border-radius: 20px;
            border: 1px solid rgba(0,245,255,0.2);
            margin-bottom: 30px;
            backdrop-filter: blur(10px);
        }
        .main-header h1 {
            color: #00f5ff;
            font-size: 2.5rem;
            text-shadow: 0 0 30px rgba(0,245,255,0.3);
            margin: 0;
        }
        .main-header p {
            color: rgba(255,255,255,0.6);
            font-size: 1.1rem;
            margin: 10px 0 0 0;
        }
        .metric-card {
            background: rgba(255, 255, 255, 0.03);
            backdrop-filter: blur(10px);
            border-radius: 15px;
            padding: 20px;
            border: 1px solid rgba(255, 255, 255, 0.08);
            box-shadow: 0 8px 32px rgba(0,0,0,0.3);
            transition: all 0.3s ease;
            height: 100%;
        }
        .metric-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 12px 48px rgba(0,245,255,0.1);
            border-color: rgba(0,245,255,0.2);
        }
        .metric-value {
            font-size: 2.2rem;
            font-weight: bold;
            background: linear-gradient(135deg, #00f5ff, #00b4ff);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
        }
        .metric-value-negative {
            font-size: 2.2rem;
            font-weight: bold;
            background: linear-gradient(135deg, #ff6b6b, #ff3366);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
        }
        .metric-label {
            color: rgba(255,255,255,0.6);
            font-size: 0.85rem;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-top: 5px;
        }
        .metric-sub {
            color: rgba(255,255,255,0.4);
            font-size: 0.8rem;
            margin-top: 8px;
        }
        .glass-card {
            background: rgba(255, 255, 255, 0.03);
            backdrop-filter: blur(10px);
            border-radius: 15px;
            padding: 20px;
            border: 1px solid rgba(255, 255, 255, 0.08);
            box-shadow: 0 8px 32px rgba(0,0,0,0.2);
            transition: all 0.3s ease;
        }
        .glass-card:hover {
            border-color: rgba(0,245,255,0.15);
        }
        .glass-card h3 {
            color: #00f5ff;
            font-size: 1.1rem;
            margin-bottom: 15px;
            text-shadow: 0 0 20px rgba(0,245,255,0.1);
        }
        .stButton > button {
            background: linear-gradient(135deg, #00f5ff, #00b4ff);
            color: white;
            border: none;
            border-radius: 10px;
            padding: 10px 30px;
            font-weight: bold;
            transition: all 0.3s ease;
            box-shadow: 0 4px 15px rgba(0, 245, 255, 0.2);
            width: 100%;
        }
        .stButton > button:hover {
            transform: scale(1.02);
            box-shadow: 0 8px 25px rgba(0, 245, 255, 0.4);
        }
        .stSelectbox, .stDateInput {
            background: rgba(255,255,255,0.05);
            border-radius: 10px;
        }
        .st-emotion-cache-1r6slb0 {
            background: rgba(255, 255, 255, 0.03);
        }
        h2, h3 {
            color: #00f5ff !important;
        }
        .stat-box {
            background: rgba(0,245,255,0.05);
            padding: 15px;
            border-radius: 10px;
            border-left: 3px solid #00f5ff;
            margin: 5px 0;
        }
        .stat-box-red {
            background: rgba(255,107,107,0.05);
            padding: 15px;
            border-radius: 10px;
            border-left: 3px solid #ff6b6b;
            margin: 5px 0;
        }
        .profit-positive {
            color: #00ff88;
            font-weight: bold;
        }
        .profit-negative {
            color: #ff6b6b;
            font-weight: bold;
        }
        .dataframe {
            background: rgba(255,255,255,0.02);
            border-radius: 10px;
        }
        .st-emotion-cache-16idsys {
            background: rgba(255,255,255,0.02);
        }
        .download-btn {
            background: linear-gradient(135deg, #00ff88, #00cc66) !important;
            box-shadow: 0 4px 15px rgba(0,255,136,0.2) !important;
        }
        .download-btn:hover {
            box-shadow: 0 8px 25px rgba(0,255,136,0.4) !important;
        }
        .sidebar-content {
            padding: 10px 0;
        }
        .sidebar-label {
            color: rgba(255,255,255,0.5);
            font-size: 0.8rem;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-top: 15px;
            margin-bottom: 5px;
        }
        hr {
            border-color: rgba(255,255,255,0.05);
            margin: 15px 0;
        }
        .footer {
            text-align: center;
            color: rgba(255,255,255,0.2);
            font-size: 0.7rem;
            padding: 20px 0;
            border-top: 1px solid rgba(255,255,255,0.05);
            margin-top: 30px;
        }
    </style>
""", unsafe_allow_html=True)

# Генерация тестовых данных для демонстрации
@st.cache_data(ttl=3600)
def generate_test_data():
    """Генерирует тестовые данные для демонстрации"""
    np.random.seed(42)
    
    # Данные для общей таблицы (Биллинг)
    dates = pd.date_range('2026-01-01', '2026-06-21', freq='D')
    regions = ['Центральный', 'Северо-Западный', 'Южный', 'Приволжский', 'Уральский', 'Сибирский']
    transport_types = ['Грузовик', 'Фургон', 'Рефрижератор', 'Тент']
    stores = ['Магнит А', 'Магнит Б', 'Магнит В', 'Магнит Г', 'Магнит Д', 'Магнит Е', 'Магнит Ж']
    
    billing_data = []
    hired_data = []
    
    for date in dates:
        # Для каждой даты генерируем от 2 до 6 записей
        for _ in range(np.random.randint(2, 7)):
            region = np.random.choice(regions)
            transport = np.random.choice(transport_types)
            store = np.random.choice(stores)
            quantity = np.random.randint(80, 500)
            
            # Биллинг (доход от клиента)
            billing_data.append({
                'Дата': date,
                'Регион': region,
                'Тип ТС': transport,
                'Магазин': store,
                'Кол-во шт': quantity,
                'Тип': 'Биллинг'
            })
            
            # Наём (расходы на исполнителей)
            hired_quantity = quantity * np.random.uniform(0.7, 0.95)
            hired_data.append({
                'Дата': date,
                'Регион': region,
                'Тип ТС': transport,
                'Магазин': store,
                'Кол-во шт': int(hired_quantity),
                'Тип': 'Наём'
            })
    
    billing_df = pd.DataFrame(billing_data)
    hired_df = pd.DataFrame(hired_data)
    
    return billing_df, hired_df

# Функция для загрузки данных из Google Sheets
@st.cache_data(ttl=3600)
def load_from_google_sheets():
    """Пытается загрузить данные из Google Sheets, при ошибке генерирует тестовые"""
    try:
        # Пытаемся импортировать gspread
        import gspread
        from oauth2client.service_account import ServiceAccountCredentials
        
        # Проверяем наличие secrets
        if 'google' in st.secrets:
            scope = ['https://spreadsheets.google.com/feeds',
                    'https://www.googleapis.com/auth/drive']
            
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
            
            # Загружаем общую таблицу (Биллинг)
            billing_worksheet = sheet.worksheet('Общ таблица')
            billing_data = billing_worksheet.get_all_values()
            
            if len(billing_data) > 1:
                headers = billing_data[0]
                rows = billing_data[1:]
                billing_df = pd.DataFrame(rows, columns=headers)
                
                # Преобразуем колонки
                billing_df['Дата'] = pd.to_datetime(billing_df['Дата'])
                billing_df['Кол-во шт'] = pd.to_numeric(billing_df['Кол-во шт'], errors='coerce')
                billing_df['Тип'] = 'Биллинг'
                
                # Переименовываем колонки для соответствия
                col_map = {
                    'B:B': 'Дата',
                    'D:D': 'Кол-во шт',
                    'O:O': 'Сумма'
                }
                
                # Загружаем сводную по дням (Наём)
                try:
                    hired_worksheet = sheet.worksheet('Сводная по дням')
                    hired_data = hired_worksheet.get_all_values()
                    
                    if len(hired_data) > 1:
                        hired_headers = hired_data[0]
                        hired_rows = hired_data[1:]
                        hired_df = pd.DataFrame(hired_rows, columns=hired_headers)
                        hired_df['Дата'] = pd.to_datetime(hired_df['Дата'])
                        hired_df['Кол-во шт'] = pd.to_numeric(hired_df['Кол-во шт'], errors='coerce')
                        hired_df['Тип'] = 'Наём'
                    else:
                        hired_df = pd.DataFrame()
                except:
                    hired_df = pd.DataFrame()
                
                return billing_df, hired_df
            else:
                st.warning("Таблица пуста, используются тестовые данные")
                return generate_test_data()
        else:
            st.info("Secrets не настроены, используются тестовые данные")
            return generate_test_data()
            
    except Exception as e:
        st.warning(f"Не удалось загрузить данные из Google Sheets: {e}")
        st.info("Используются тестовые данные для демонстрации")
        return generate_test_data()

# Функция для создания Excel выгрузки по найму
def create_hired_excel(hired_df, filename="naym_analytics.xlsx"):
    """Создает Excel файл с данными по найму"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Основные данные по найму
        hired_df.to_excel(writer, sheet_name='Наём_данные', index=False)
        
        # Сводка по дням
        daily_hired = hired_df.groupby('Дата').agg({
            'Кол-во шт': 'sum'
        }).reset_index()
        daily_hired.to_excel(writer, sheet_name='Наём_по_дням', index=False)
        
        # Сводка по регионам
        region_hired = hired_df.groupby('Регион').agg({
            'Кол-во шт': 'sum'
        }).reset_index()
        region_hired.to_excel(writer, sheet_name='Наём_по_регионам', index=False)
        
        # Сводка по типам ТС
        transport_hired = hired_df.groupby('Тип ТС').agg({
            'Кол-во шт': 'sum'
        }).reset_index()
        transport_hired.to_excel(writer, sheet_name='Наём_по_ТС', index=False)
        
        # Детализация по магазинам
        store_hired = hired_df.groupby(['Дата', 'Регион', 'Магазин']).agg({
            'Кол-во шт': 'sum'
        }).reset_index()
        store_hired.to_excel(writer, sheet_name='Наём_по_магазинам', index=False)
        
        # Сводная таблица
        pivot_table = pd.pivot_table(
            hired_df,
            values='Кол-во шт',
            index=['Дата'],
            columns=['Регион'],
            aggfunc='sum',
            fill_value=0
        )
        pivot_table.to_excel(writer, sheet_name='Сводная_по_регионам')
    
    output.seek(0)
    return output

# Основная функция дашборда
def main():
    # Заголовок
    st.markdown("""
        <div class="main-header">
            <h1>🚛 Аналитика логистики</h1>
            <p>Сравнительный анализ доходов от клиента и расходов на наём</p>
        </div>
    """, unsafe_allow_html=True)
    
    # Загрузка данных
    with st.spinner("Загрузка данных..."):
        billing_df, hired_df = load_from_google_sheets()
    
    # Если данные пустые, генерируем тестовые
    if billing_df.empty:
        billing_df, hired_df = generate_test_data()
    
    # Преобразование дат
    billing_df['Дата'] = pd.to_datetime(billing_df['Дата'])
    if not hired_df.empty:
        hired_df['Дата'] = pd.to_datetime(hired_df['Дата'])
    
    # Боковая панель с фильтрами
    with st.sidebar:
        st.markdown("""
            <div class="sidebar-content">
                <h3 style="color: #00f5ff; margin-bottom: 20px;">🎯 Фильтры</h3>
            </div>
        """, unsafe_allow_html=True)
        
        # Выбор периода
        min_date = min(billing_df['Дата'].min().date(), hired_df['Дата'].min().date() if not hired_df.empty else billing_df['Дата'].min().date())
        max_date = max(billing_df['Дата'].max().date(), hired_df['Дата'].max().date() if not hired_df.empty else billing_df['Дата'].max().date())
        
        date_range = st.date_input(
            "📅 Период",
            [min_date, max_date],
            min_value=min_date,
            max_value=max_date
        )
        
        if len(date_range) == 2:
            start_date, end_date = date_range
            billing_mask = (billing_df['Дата'].dt.date >= start_date) & (billing_df['Дата'].dt.date <= end_date)
            billing_filtered = billing_df.loc[billing_mask]
            
            if not hired_df.empty:
                hired_mask = (hired_df['Дата'].dt.date >= start_date) & (hired_df['Дата'].dt.date <= end_date)
                hired_filtered = hired_df.loc[hired_mask]
            else:
                hired_filtered = pd.DataFrame()
        else:
            billing_filtered = billing_df
            hired_filtered = hired_df
        
        st.markdown("<hr>", unsafe_allow_html=True)
        
        # Фильтр по регионам
        all_regions = sorted(set(billing_filtered['Регион'].unique()).union(
            set(hired_filtered['Регион'].unique()) if not hired_filtered.empty else set()
        ))
        
        regions = st.multiselect(
            "📍 Регионы",
            options=all_regions,
            default=all_regions
        )
        
        if regions:
            billing_filtered = billing_filtered[billing_filtered['Регион'].isin(regions)]
            if not hired_filtered.empty:
                hired_filtered = hired_filtered[hired_filtered['Регион'].isin(regions)]
        
        st.markdown("<hr>", unsafe_allow_html=True)
        
        # Фильтр по типам ТС
        all_transport = sorted(set(billing_filtered['Тип ТС'].unique()).union(
            set(hired_filtered['Тип ТС'].unique()) if not hired_filtered.empty else set()
        ))
        
        transport_types = st.multiselect(
            "🚚 Тип ТС",
            options=all_transport,
            default=all_transport
        )
        
        if transport_types:
            billing_filtered = billing_filtered[billing_filtered['Тип ТС'].isin(transport_types)]
            if not hired_filtered.empty:
                hired_filtered = hired_filtered[hired_filtered['Тип ТС'].isin(transport_types)]
        
        st.markdown("<hr>", unsafe_allow_html=True)
        
        # Кнопка обновления
        if st.button("🔄 Обновить данные", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
        
        st.markdown("""
            <div style="margin-top: 20px; padding: 10px; background: rgba(0,245,255,0.05); border-radius: 10px;">
                <p style="color: rgba(255,255,255,0.4); font-size: 0.7rem; text-align: center;">
                    Данные обновлены: {}
                </p>
            </div>
        """.format(datetime.now().strftime("%d.%m.%Y %H:%M")), unsafe_allow_html=True)
    
    # Основной контент
    # ===== КЛЮЧЕВЫЕ МЕТРИКИ =====
    st.markdown("## 📊 Ключевые показатели")
    
    col1, col2, col3, col4 = st.columns(4)
    
    # Метрики по биллингу
    billing_quantity = billing_filtered['Кол-во шт'].sum()
    billing_days = billing_filtered['Дата'].nunique()
    billing_avg_per_day = billing_quantity / billing_days if billing_days > 0 else 0
    
    # Метрики по найму
    hired_quantity = hired_filtered['Кол-во шт'].sum() if not hired_filtered.empty else 0
    hired_days = hired_filtered['Дата'].nunique() if not hired_filtered.empty else 0
    hired_avg_per_day = hired_quantity / hired_days if hired_days > 0 else 0
    
    # Разница
    diff_quantity = billing_quantity - hired_quantity
    diff_percent = (diff_quantity / billing_quantity * 100) if billing_quantity > 0 else 0
    
    with col1:
        st.markdown(f"""
            <div class="metric-card">
                <div style="display: flex; align-items: center; gap: 10px;">
                    <span style="font-size: 2rem;">💰</span>
                    <div>
                        <div class="metric-value">{billing_quantity:,.0f}</div>
                        <div class="metric-label">Биллинг (доход от клиента)</div>
                        <div class="metric-sub">в среднем {billing_avg_per_day:,.0f} шт/день</div>
                    </div>
                </div>
            </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
            <div class="metric-card">
                <div style="display: flex; align-items: center; gap: 10px;">
                    <span style="font-size: 2rem;">💳</span>
                    <div>
                        <div class="metric-value" style="background: linear-gradient(135deg, #ff6b6b, #ff3366); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;">{hired_quantity:,.0f}</div>
                        <div class="metric-label">Наём (расходы)</div>
                        <div class="metric-sub">в среднем {hired_avg_per_day:,.0f} шт/день</div>
                    </div>
                </div>
            </div>
        """, unsafe_allow_html=True)
    
    with col3:
        profit_color = "metric-value" if diff_quantity >= 0 else "metric-value-negative"
        st.markdown(f"""
            <div class="metric-card">
                <div style="display: flex; align-items: center; gap: 10px;">
                    <span style="font-size: 2rem;">📈</span>
                    <div>
                        <div class="{profit_color}">{diff_quantity:+,.0f}</div>
                        <div class="metric-label">Чистая прибыль (шт)</div>
                        <div class="metric-sub" style="color: {'#00ff88' if diff_quantity >= 0 else '#ff6b6b'};">
                            {diff_percent:+.1f}% от биллинга
                        </div>
                    </div>
                </div>
            </div>
        """, unsafe_allow_html=True)
    
    with col4:
        efficiency = (billing_quantity / hired_quantity) if hired_quantity > 0 else 0
        st.markdown(f"""
            <div class="metric-card">
                <div style="display: flex; align-items: center; gap: 10px;">
                    <span style="font-size: 2rem;">⚡</span>
                    <div>
                        <div class="metric-value" style="background: linear-gradient(135deg, #ffd93d, #ff9a3d); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;">{efficiency:.2f}x</div>
                        <div class="metric-label">Эффективность</div>
                        <div class="metric-sub">Биллинг / Наём</div>
                    </div>
                </div>
            </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # ===== ГРАФИКИ =====
    st.markdown("## 📈 Аналитика")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.markdown("#### 📊 Динамика по дням")
        
        # Объединяем данные для графика
        daily_billing = billing_filtered.groupby('Дата')['Кол-во шт'].sum().reset_index()
        daily_billing['Тип'] = 'Биллинг'
        
        if not hired_filtered.empty:
            daily_hired = hired_filtered.groupby('Дата')['Кол-во шт'].sum().reset_index()
            daily_hired['Тип'] = 'Наём'
            daily_combined = pd.concat([daily_billing, daily_hired])
        else:
            daily_combined = daily_billing
        
        fig = px.line(
            daily_combined,
            x='Дата',
            y='Кол-во шт',
            color='Тип',
            color_discrete_map={'Биллинг': '#00f5ff', 'Наём': '#ff6b6b'},
            markers=True
        )
        fig.update_traces(marker=dict(size=6))
        fig.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white',
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(l=0, r=0, t=20, b=0),
            height=350,
            xaxis=dict(gridcolor='rgba(255,255,255,0.05)', title=""),
            yaxis=dict(gridcolor='rgba(255,255,255,0.05)', title="Кол-во шт")
        )
        st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.markdown("#### 🗺️ Распределение по регионам")
        
        # Данные по регионам
        billing_regions = billing_filtered.groupby('Регион')['Кол-во шт'].sum().reset_index()
        billing_regions['Тип'] = 'Биллинг'
        
        if not hired_filtered.empty:
            hired_regions = hired_filtered.groupby('Регион')['Кол-во шт'].sum().reset_index()
            hired_regions['Тип'] = 'Наём'
            regions_combined = pd.concat([billing_regions, hired_regions])
        else:
            regions_combined = billing_regions
        
        fig = px.bar(
            regions_combined,
            x='Регион',
            y='Кол-во шт',
            color='Тип',
            color_discrete_map={'Биллинг': '#00f5ff', 'Наём': '#ff6b6b'},
            barmode='group',
            text='Кол-во шт'
        )
        fig.update_traces(textposition='outside', textfont_color='white')
        fig.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white',
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(l=0, r=0, t=20, b=0),
            height=350,
            xaxis=dict(gridcolor='rgba(255,255,255,0.05)', title=""),
            yaxis=dict(gridcolor='rgba(255,255,255,0.05)', title="Кол-во шт")
        )
        st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    # ===== ДОПОЛНИТЕЛЬНЫЕ ГРАФИКИ =====
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.markdown("#### 🚚 Распределение по типам ТС")
        
        billing_transport = billing_filtered.groupby('Тип ТС')['Кол-во шт'].sum().reset_index()
        
        fig = px.pie(
            billing_transport,
            values='Кол-во шт',
            names='Тип ТС',
            color_discrete_sequence=px.colors.sequential.Teal_r,
            hole=0.4
        )
        fig.update_traces(textposition='inside', textinfo='percent+label', textfont_color='white')
        fig.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white',
            showlegend=False,
            margin=dict(l=0, r=0, t=10, b=0),
            height=300
        )
        st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.markdown("#### 📊 Соотношение Биллинг vs Наём")
        
        comparison_data = pd.DataFrame({
            'Категория': ['Биллинг', 'Наём'],
            'Кол-во шт': [billing_quantity, hired_quantity]
        })
        
        fig = px.bar(
            comparison_data,
            x='Категория',
            y='Кол-во шт',
            color='Категория',
            color_discrete_map={'Биллинг': '#00f5ff', 'Наём': '#ff6b6b'},
            text='Кол-во шт'
        )
        fig.update_traces(textposition='outside', textfont_color='white')
        fig.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white',
            showlegend=False,
            margin=dict(l=0, r=0, t=10, b=0),
            height=300,
            xaxis=dict(gridcolor='rgba(255,255,255,0.05)', title=""),
            yaxis=dict(gridcolor='rgba(255,255,255,0.05)', title="Кол-во шт")
        )
        st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    # ===== ДЕТАЛЬНАЯ ТАБЛИЦА =====
    st.markdown("## 📋 Детализация данных")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        # Выбор даты для детализации
        detail_date = st.date_input(
            "📅 Выберите дату",
            value=billing_filtered['Дата'].max().date() if not billing_filtered.empty else datetime.now().date(),
            min_value=billing_filtered['Дата'].min().date() if not billing_filtered.empty else datetime.now().date() - timedelta(days=30),
            max_value=billing_filtered['Дата'].max().date() if not billing_filtered.empty else datetime.now().date()
        )
    
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        # Кнопка выгрузки Excel
        if not hired_filtered.empty:
            if st.button("📥 Выгрузить данные по найму в Excel", use_container_width=True):
                excel_file = create_hired_excel(hired_filtered)
                b64 = base64.b64encode(excel_file.getvalue()).decode()
                href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="naym_analytics.xlsx" style="text-decoration: none;">📊 Скачать Excel</a>'
                st.markdown(href, unsafe_allow_html=True)
        else:
            st.info("Нет данных по найму для выгрузки")
    
    # Детальная таблица по выбранной дате
    detail_billing = billing_filtered[billing_filtered['Дата'].dt.date == detail_date]
    
    if not detail_billing.empty:
        # Группировка данных
        detail_table = detail_billing.groupby(['Регион', 'Тип ТС']).agg({
            'Магазин': lambda x: ', '.join(x.unique()),
            'Кол-во шт': 'sum'
        }).reset_index()
        
        # Добавляем информацию о найме для этой даты
        if not hired_filtered.empty:
            detail_hired = hired_filtered[hired_filtered['Дата'].dt.date == detail_date]
            if not detail_hired.empty:
                hired_by_region = detail_hired.groupby('Регион')['Кол-во шт'].sum().reset_index()
                hired_by_region.columns = ['Регион', 'Наём_шт']
                detail_table = detail_table.merge(hired_by_region, on='Регион', how='left')
                detail_table['Наём_шт'] = detail_table['Наём_шт'].fillna(0)
            else:
                detail_table['Наём_шт'] = 0
        else:
            detail_table['Наём_шт'] = 0
        
        # Расчет разницы
        detail_table['Разница'] = detail_table['Кол-во шт'] - detail_table['Наём_шт']
        
        st.dataframe(
            detail_table,
            use_container_width=True,
            column_config={
                'Регион': '📍 Регион',
                'Тип ТС': '🚚 Тип ТС',
                'Магазин': '🏪 Магазины',
                'Кол-во шт': st.column_config.NumberColumn('📦 Биллинг (шт)', format='%d'),
                'Наём_шт': st.column_config.NumberColumn('👷 Наём (шт)', format='%d'),
                'Разница': st.column_config.NumberColumn('📈 Разница (шт)', format='%d')
            },
            hide_index=True
        )
    else:
        st.info("📭 Нет данных за выбранную дату")
    
    # ===== СТАТИСТИКА ПО РЕГИОНАМ =====
    st.markdown("## 📊 Статистика по регионам")
    
    # Создаем сводную таблицу
    region_stats = billing_filtered.groupby('Регион').agg({
        'Кол-во шт': 'sum',
        'Магазин': 'nunique'
    }).reset_index()
    region_stats.columns = ['Регион', 'Биллинг_шт', 'Кол-во_точек']
    
    if not hired_filtered.empty:
        hired_stats = hired_filtered.groupby('Регион')['Кол-во шт'].sum().reset_index()
        hired_stats.columns = ['Регион', 'Наём_шт']
        region_stats = region_stats.merge(hired_stats, on='Регион', how='left')
        region_stats['Наём_шт'] = region_stats['Наём_шт'].fillna(0)
    else:
        region_stats['Наём_шт'] = 0
    
    region_stats['Разница'] = region_stats['Биллинг_шт'] - region_stats['Наём_шт']
    region_stats['Эффективность'] = (region_stats['Биллинг_шт'] / region_stats['Наём_шт']).round(2)
    region_stats['Эффективность'] = region_stats['Эффективность'].replace([float('inf'), -float('inf')], 0)
    
    st.dataframe(
        region_stats,
        use_container_width=True,
        column_config={
            'Регион': '📍 Регион',
            'Биллинг_шт': st.column_config.NumberColumn('📦 Биллинг (шт)', format='%d'),
            'Наём_шт': st.column_config.NumberColumn('👷 Наём (шт)', format='%d'),
            'Разница': st.column_config.NumberColumn('📈 Прибыль (шт)', format='%d'),
            'Кол-во_точек': st.column_config.NumberColumn('🏪 Кол-во точек', format='%d'),
            'Эффективность': st.column_config.NumberColumn('⚡ Эффективность', format='%.2f x')
        },
        hide_index=True
    )
    
    # ===== ДОПОЛНИТЕЛЬНАЯ ИНФОРМАЦИЯ =====
    with st.expander("📌 О дашборде", expanded=False):
        st.markdown("""
            ### Как это работает:
            
            **📊 Основная логика:**
            - **Биллинг** - это доход, который вы получаете от клиента (Магнит)
            - **Наём** - это расходы на исполнителей (ваши затраты)
            - **Чистая прибыль** = Биллинг - Наём (в штуках)
            
            **🎯 Что показывает дашборд:**
            1. Ежедневную динамику доходов и расходов
            2. Распределение по регионам и типам ТС
            3. Эффективность работы в каждом регионе
            4. Детализацию по дням и магазинам
            
            **📥 Выгрузка Excel:**
            - Нажмите кнопку "Выгрузить данные по найму в Excel"
            - Получите детальный отчет по расходам на наём
            - Включает сводки по дням, регионам, ТС и магазинам
        """)
    
    # Footer
    st.markdown("""
        <div class="footer">
            🚛 Аналитический дашборд логистики | Биллинг vs Наём
        </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
