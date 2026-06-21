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
            padding: 25px;
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
        h2, h3 {
            color: #00f5ff !important;
        }
        .profit-positive {
            color: #00ff88;
            font-weight: bold;
        }
        .profit-negative {
            color: #ff6b6b;
            font-weight: bold;
        }
        .footer {
            text-align: center;
            color: rgba(255,255,255,0.2);
            font-size: 0.7rem;
            padding: 20px 0;
            border-top: 1px solid rgba(255,255,255,0.05);
            margin-top: 30px;
        }
        .stDataFrame {
            background: rgba(255,255,255,0.02);
            border-radius: 10px;
        }
        .info-box {
            background: rgba(0,245,255,0.05);
            border-left: 3px solid #00f5ff;
            padding: 15px;
            border-radius: 5px;
            margin: 10px 0;
        }
        .warning-box {
            background: rgba(255,107,107,0.05);
            border-left: 3px solid #ff6b6b;
            padding: 15px;
            border-radius: 5px;
            margin: 10px 0;
        }
    </style>
""", unsafe_allow_html=True)

# Функция для определения структуры данных
def detect_data_structure(df):
    """Определяет структуру данных и возвращает маппинг колонок"""
    df_lower = {col.lower(): col for col in df.columns}
    
    # Маппинг колонок
    mapping = {
        'date': None,
        'quantity': None,
        'amount': None,
        'region': None,
        'transport': None,
        'store': None,
        'type': None
    }
    
    # Поиск колонок
    date_keywords = ['дата', 'date', 'день', 'day', 'период', 'period', 'b:b']
    quantity_keywords = ['кол-во', 'количество', 'quantity', 'qty', 'шт', 'd:d']
    amount_keywords = ['сумма', 'amount', 'total', 'стоимость', 'o:o']
    region_keywords = ['регион', 'region', 'область', 'area']
    transport_keywords = ['тип', 'тс', 'transport', 'авто', 'car', 'type']
    store_keywords = ['магазин', 'store', 'точка', 'shop', 'клиент']
    
    for col in df.columns:
        col_lower = col.lower()
        if mapping['date'] is None and any(kw in col_lower for kw in date_keywords):
            mapping['date'] = col
        elif mapping['quantity'] is None and any(kw in col_lower for kw in quantity_keywords):
            mapping['quantity'] = col
        elif mapping['amount'] is None and any(kw in col_lower for kw in amount_keywords):
            mapping['amount'] = col
        elif mapping['region'] is None and any(kw in col_lower for kw in region_keywords):
            mapping['region'] = col
        elif mapping['transport'] is None and any(kw in col_lower for kw in transport_keywords):
            mapping['transport'] = col
        elif mapping['store'] is None and any(kw in col_lower for kw in store_keywords):
            mapping['store'] = col
    
    return mapping

# Генерация тестовых данных
def generate_test_data():
    """Генерирует тестовые данные для демонстрации"""
    np.random.seed(42)
    
    dates = pd.date_range('2026-01-01', '2026-06-21', freq='D')
    regions = ['Центральный', 'Северо-Западный', 'Южный', 'Приволжский', 'Уральский', 'Сибирский']
    transport_types = ['Грузовик', 'Фургон', 'Рефрижератор', 'Тент']
    stores = ['Магнит А', 'Магнит Б', 'Магнит В', 'Магнит Г', 'Магнит Д', 'Магнит Е']
    
    billing_data = []
    hired_data = []
    
    for date in dates:
        for _ in range(np.random.randint(2, 6)):
            region = np.random.choice(regions)
            transport = np.random.choice(transport_types)
            store = np.random.choice(stores)
            quantity = np.random.randint(80, 500)
            
            # Биллинг
            billing_data.append({
                'Дата': date,
                'Регион': region,
                'Тип ТС': transport,
                'Магазин': store,
                'Кол-во шт': quantity,
                'Тип': 'Биллинг'
            })
            
            # Наём
            hired_data.append({
                'Дата': date,
                'Регион': region,
                'Тип ТС': transport,
                'Магазин': store,
                'Кол-во шт': int(quantity * np.random.uniform(0.7, 0.95)),
                'Тип': 'Наём'
            })
    
    return pd.DataFrame(billing_data), pd.DataFrame(hired_data)

# Функция для загрузки данных
@st.cache_data(ttl=3600)
def load_data():
    """Загрузка данных из различных источников"""
    try:
        # Проверяем, есть ли файлы в сессии
        if 'billing_df' in st.session_state and 'hired_df' in st.session_state:
            if not st.session_state.billing_df.empty and not st.session_state.hired_df.empty:
                return st.session_state.billing_df, st.session_state.hired_df
        
        # Пытаемся загрузить из Google Sheets
        try:
            import gspread
            from oauth2client.service_account import ServiceAccountCredentials
            
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
                
                # Загружаем общую таблицу
                billing_ws = sheet.worksheet('Общ таблица')
                billing_data = billing_ws.get_all_values()
                
                if len(billing_data) > 1:
                    headers = billing_data[0]
                    rows = billing_data[1:]
                    billing_df = pd.DataFrame(rows, columns=headers)
                    
                    # Определяем структуру данных
                    mapping = detect_data_structure(billing_df)
                    
                    # Преобразуем данные
                    if mapping['date']:
                        billing_df['Дата'] = pd.to_datetime(billing_df[mapping['date']], errors='coerce')
                    if mapping['quantity']:
                        billing_df['Кол-во шт'] = pd.to_numeric(billing_df[mapping['quantity']], errors='coerce')
                    if mapping['amount']:
                        billing_df['Сумма'] = pd.to_numeric(billing_df[mapping['amount']], errors='coerce')
                    if mapping['region']:
                        billing_df['Регион'] = billing_df[mapping['region']]
                    if mapping['transport']:
                        billing_df['Тип ТС'] = billing_df[mapping['transport']]
                    if mapping['store']:
                        billing_df['Магазин'] = billing_df[mapping['store']]
                    
                    billing_df['Тип'] = 'Биллинг'
                    
                    # Загружаем сводную по дням (наём)
                    try:
                        hired_ws = sheet.worksheet('Сводная по дням')
                        hired_data = hired_ws.get_all_values()
                        
                        if len(hired_data) > 1:
                            hired_headers = hired_data[0]
                            hired_rows = hired_data[1:]
                            hired_df = pd.DataFrame(hired_rows, columns=hired_headers)
                            
                            mapping_hired = detect_data_structure(hired_df)
                            
                            if mapping_hired['date']:
                                hired_df['Дата'] = pd.to_datetime(hired_df[mapping_hired['date']], errors='coerce')
                            if mapping_hired['quantity']:
                                hired_df['Кол-во шт'] = pd.to_numeric(hired_df[mapping_hired['quantity']], errors='coerce')
                            if mapping_hired['region']:
                                hired_df['Регион'] = hired_df[mapping_hired['region']]
                            if mapping_hired['transport']:
                                hired_df['Тип ТС'] = hired_df[mapping_hired['transport']]
                            if mapping_hired['store']:
                                hired_df['Магазин'] = hired_df[mapping_hired['store']]
                            
                            hired_df['Тип'] = 'Наём'
                            
                            # Сохраняем в сессию
                            st.session_state.billing_df = billing_df
                            st.session_state.hired_df = hired_df
                            
                            return billing_df, hired_df
                    except Exception as e:
                        st.warning(f"Не удалось загрузить данные по найму: {e}")
                    
                    return billing_df, pd.DataFrame()
        except Exception as e:
            st.warning(f"Не удалось загрузить данные из Google Sheets: {e}")
        
        # Если не удалось загрузить, генерируем тестовые
        billing_df, hired_df = generate_test_data()
        st.session_state.billing_df = billing_df
        st.session_state.hired_df = hired_df
        return billing_df, hired_df
        
    except Exception as e:
        st.error(f"Ошибка загрузки данных: {e}")
        return generate_test_data()

# Функция для создания Excel выгрузки
def create_excel_export(billing_df, hired_df, filename="analytics_export.xlsx"):
    """Создает Excel файл с данными"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Данные по биллингу
        billing_export = billing_df[['Дата', 'Регион', 'Тип ТС', 'Магазин', 'Кол-во шт']].copy()
        billing_export.to_excel(writer, sheet_name='Биллинг', index=False)
        
        # Данные по найму
        if not hired_df.empty:
            hired_export = hired_df[['Дата', 'Регион', 'Тип ТС', 'Магазин', 'Кол-во шт']].copy()
            hired_export.to_excel(writer, sheet_name='Наём', index=False)
            
            # Сравнение
            daily_billing = billing_df.groupby('Дата')['Кол-во шт'].sum().reset_index()
            daily_billing.columns = ['Дата', 'Биллинг_шт']
            
            daily_hired = hired_df.groupby('Дата')['Кол-во шт'].sum().reset_index()
            daily_hired.columns = ['Дата', 'Наём_шт']
            
            comparison = pd.merge(daily_billing, daily_hired, on='Дата', how='outer').fillna(0)
            comparison['Разница'] = comparison['Биллинг_шт'] - comparison['Наём_шт']
            comparison.to_excel(writer, sheet_name='Сравнение_по_дням', index=False)
            
            # Сводка по регионам
            region_billing = billing_df.groupby('Регион')['Кол-во шт'].sum().reset_index()
            region_billing.columns = ['Регион', 'Биллинг_шт']
            
            region_hired = hired_df.groupby('Регион')['Кол-во шт'].sum().reset_index()
            region_hired.columns = ['Регион', 'Наём_шт']
            
            region_summary = pd.merge(region_billing, region_hired, on='Регион', how='outer').fillna(0)
            region_summary['Разница'] = region_summary['Биллинг_шт'] - region_summary['Наём_шт']
            region_summary.to_excel(writer, sheet_name='Сводка_по_регионам', index=False)
    
    output.seek(0)
    return output

# Основная функция
def main():
    # Заголовок
    st.markdown("""
        <div class="main-header">
            <h1>🚛 Аналитика логистики</h1>
            <p>📊 Доход от клиента (Биллинг) vs Расходы на наём</p>
        </div>
    """, unsafe_allow_html=True)
    
    # Загрузка данных
    with st.spinner("🔄 Загрузка данных..."):
        billing_df, hired_df = load_data()
    
    # Очистка данных
    billing_df = billing_df.dropna(subset=['Кол-во шт'])
    billing_df['Дата'] = pd.to_datetime(billing_df['Дата'])
    
    if not hired_df.empty:
        hired_df = hired_df.dropna(subset=['Кол-во шт'])
        hired_df['Дата'] = pd.to_datetime(hired_df['Дата'])
    
    # Боковая панель
    with st.sidebar:
        st.markdown("""
            <div style="padding: 10px 0;">
                <h3 style="color: #00f5ff; margin-bottom: 20px;">🎯 Фильтры</h3>
            </div>
        """, unsafe_allow_html=True)
        
        # Выбор периода
        min_date = min(billing_df['Дата'].min().date(), 
                      hired_df['Дата'].min().date() if not hired_df.empty else billing_df['Дата'].min().date())
        max_date = max(billing_df['Дата'].max().date(), 
                      hired_df['Дата'].max().date() if not hired_df.empty else billing_df['Дата'].max().date())
        
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
        
        # Фильтр по регионам
        all_regions = sorted(set(billing_filtered['Регион'].unique()).union(
            set(hired_filtered['Регион'].unique()) if not hired_filtered.empty else set()
        ))
        
        regions = st.multiselect(
            "📍 Регионы",
            options=all_regions,
            default=all_regions[:3] if len(all_regions) > 3 else all_regions
        )
        
        if regions:
            billing_filtered = billing_filtered[billing_filtered['Регион'].isin(regions)]
            if not hired_filtered.empty:
                hired_filtered = hired_filtered[hired_filtered['Регион'].isin(regions)]
        
        st.markdown("<hr>", unsafe_allow_html=True)
        
        # Информация о данных
        st.markdown(f"""
            <div style="background: rgba(0,245,255,0.03); padding: 15px; border-radius: 10px;">
                <p style="color: rgba(255,255,255,0.4); font-size: 0.8rem; margin: 0;">
                    📊 Всего записей:<br>
                    Биллинг: <b style="color: #00f5ff;">{len(billing_filtered):,}</b><br>
                    Наём: <b style="color: #ff6b6b;">{len(hired_filtered):,}</b>
                </p>
                <p style="color: rgba(255,255,255,0.3); font-size: 0.7rem; margin-top: 10px;">
                    Обновлено: {datetime.now().strftime("%d.%m.%Y %H:%M")}
                </p>
            </div>
        """, unsafe_allow_html=True)
    
    # Ключевые метрики
    st.markdown("## 📊 Ключевые показатели")
    
    col1, col2, col3, col4 = st.columns(4)
    
    billing_total = billing_filtered['Кол-во шт'].sum()
    hired_total = hired_filtered['Кол-во шт'].sum() if not hired_filtered.empty else 0
    profit = billing_total - hired_total
    efficiency = billing_total / hired_total if hired_total > 0 else 0
    
    with col1:
        st.markdown(f"""
            <div class="metric-card">
                <div style="display: flex; align-items: center; gap: 10px;">
                    <span style="font-size: 2rem;">💰</span>
                    <div>
                        <div class="metric-value">{billing_total:,.0f}</div>
                        <div class="metric-label">Биллинг (доход от клиента)</div>
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
                        <div class="metric-value" style="background: linear-gradient(135deg, #ff6b6b, #ff3366); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;">{hired_total:,.0f}</div>
                        <div class="metric-label">Наём (расходы)</div>
                    </div>
                </div>
            </div>
        """, unsafe_allow_html=True)
    
    with col3:
        profit_class = "metric-value" if profit >= 0 else "metric-value-negative"
        st.markdown(f"""
            <div class="metric-card">
                <div style="display: flex; align-items: center; gap: 10px;">
                    <span style="font-size: 2rem;">📈</span>
                    <div>
                        <div class="{profit_class}">{profit:+,.0f}</div>
                        <div class="metric-label">Чистая прибыль (шт)</div>
                    </div>
                </div>
            </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
            <div class="metric-card">
                <div style="display: flex; align-items: center; gap: 10px;">
                    <span style="font-size: 2rem;">⚡</span>
                    <div>
                        <div class="metric-value" style="background: linear-gradient(135deg, #ffd93d, #ff9a3d); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;">{efficiency:.2f}x</div>
                        <div class="metric-label">Эффективность</div>
                    </div>
                </div>
            </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Графики
    st.markdown("## 📈 Аналитика")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.markdown("#### 📊 Динамика по дням")
        
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
    
    # Детализация и выгрузка
    st.markdown("## 📋 Детализация и выгрузка")
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        detail_date = st.date_input(
            "📅 Выберите дату",
            value=billing_filtered['Дата'].max().date() if not billing_filtered.empty else datetime.now().date(),
            min_value=billing_filtered['Дата'].min().date() if not billing_filtered.empty else datetime.now().date() - timedelta(days=30),
            max_value=billing_filtered['Дата'].max().date() if not billing_filtered.empty else datetime.now().date()
        )
    
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("📥 Выгрузить в Excel", use_container_width=True):
            excel_file = create_excel_export(billing_filtered, hired_filtered)
            b64 = base64.b64encode(excel_file.getvalue()).decode()
            href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="analytics_export.xlsx" style="text-decoration: none;">📊 Скачать Excel</a>'
            st.markdown(href, unsafe_allow_html=True)
    
    # Детальная таблица
    detail_billing = billing_filtered[billing_filtered['Дата'].dt.date == detail_date]
    
    if not detail_billing.empty:
        detail_table = detail_billing.groupby(['Регион', 'Тип ТС']).agg({
            'Магазин': lambda x: ', '.join(x.unique()),
            'Кол-во шт': 'sum'
        }).reset_index()
        
        # Добавляем данные по найму
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
    
    # Сводка по регионам
    st.markdown("## 📊 Сводка по регионам")
    
    region_summary = billing_filtered.groupby('Регион').agg({
        'Кол-во шт': 'sum',
        'Магазин': 'nunique'
    }).reset_index()
    region_summary.columns = ['Регион', 'Биллинг_шт', 'Кол-во_точек']
    
    if not hired_filtered.empty:
        hired_summary = hired_filtered.groupby('Регион')['Кол-во шт'].sum().reset_index()
        hired_summary.columns = ['Регион', 'Наём_шт']
        region_summary = region_summary.merge(hired_summary, on='Регион', how='left')
        region_summary['Наём_шт'] = region_summary['Наём_шт'].fillna(0)
    else:
        region_summary['Наём_шт'] = 0
    
    region_summary['Разница'] = region_summary['Биллинг_шт'] - region_summary['Наём_шт']
    region_summary['Эффективность'] = (region_summary['Биллинг_шт'] / region_summary['Наём_шт']).round(2)
    region_summary['Эффективность'] = region_summary['Эффективность'].replace([float('inf'), -float('inf')], 0)
    
    st.dataframe(
        region_summary,
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
    
    # Footer
    st.markdown("""
        <div class="footer">
            🚛 Аналитический дашборд логистики | Биллинг vs Наём
        </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
