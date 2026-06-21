import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import io
import base64
import numpy as np
import re

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
        .dataframe {
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
    </style>
""", unsafe_allow_html=True)

# Функция для парсинга даты
def parse_date(date_str):
    """Парсинг даты в разных форматах"""
    if pd.isna(date_str):
        return None
    date_str = str(date_str).strip()
    # Пробуем разные форматы
    formats = ['%d.%m.%Y', '%Y-%m-%d', '%d.%m.%y', '%Y.%m.%d', '%d/%m/%Y']
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except:
            continue
    return None

# Функция для загрузки данных из Google Sheets
@st.cache_data(ttl=3600)
def load_data_from_sheets():
    """Загрузка данных из Google Sheets"""
    try:
        # Пытаемся импортировать gspread
        import gspread
        from oauth2client.service_account import ServiceAccountCredentials
        
        # Проверяем secrets
        if 'google' not in st.secrets:
            st.warning("Secrets не настроены. Используются тестовые данные.")
            return generate_test_data()
        
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
        
        # Загружаем данные
        billing_data = []
        hired_data = []
        
        # Пробуем получить все листы
        worksheets = sheet.worksheets()
        
        for ws in worksheets:
            title = ws.title
            data = ws.get_all_values()
            
            if len(data) < 2:
                continue
                
            headers = data[0]
            rows = data[1:]
            
            # Определяем тип таблицы по заголовкам
            if any('Биллинг' in title or 'Общ' in title or 'отгрузки' in title for title in [title]):
                # Данные биллинга
                df = pd.DataFrame(rows, columns=headers)
                billing_data.append(df)
            elif any('Сводная' in title or 'Наём' in title or 'дням' in title for title in [title]):
                # Данные найма
                df = pd.DataFrame(rows, columns=headers)
                hired_data.append(df)
        
        # Объединяем данные
        if billing_data:
            billing_df = pd.concat(billing_data, ignore_index=True)
        else:
            billing_df = pd.DataFrame()
            
        if hired_data:
            hired_df = pd.concat(hired_data, ignore_index=True)
        else:
            hired_df = pd.DataFrame()
        
        # Обрабатываем данные
        if not billing_df.empty:
            billing_df = process_billing_data(billing_df)
        if not hired_df.empty:
            hired_df = process_hired_data(hired_df)
        
        if billing_df.empty and hired_df.empty:
            st.warning("Не удалось загрузить данные. Используются тестовые данные.")
            return generate_test_data()
        
        return billing_df, hired_df
        
    except Exception as e:
        st.warning(f"Ошибка загрузки из Google Sheets: {e}")
        return generate_test_data()

def process_billing_data(df):
    """Обработка данных биллинга"""
    # Ищем нужные колонки
    for col in df.columns:
        col_lower = str(col).lower()
        if 'дата' in col_lower or 'date' in col_lower:
            df['Дата'] = df[col].apply(parse_date)
        elif 'шт' in col_lower or 'quantity' in col_lower or 'кол-во' in col_lower:
            df['Кол-во шт'] = pd.to_numeric(df[col], errors='coerce')
        elif 'регион' in col_lower or 'region' in col_lower:
            df['Регион'] = df[col]
        elif 'авто' in col_lower or 'тс' in col_lower or 'transport' in col_lower:
            df['Тип ТС'] = df[col]
        elif 'магазин' in col_lower or 'store' in col_lower:
            df['Магазин'] = df[col]
        elif 'сумма' in col_lower or 'total' in col_lower:
            df['Сумма'] = pd.to_numeric(df[col], errors='coerce')
    
    # Удаляем строки с ошибками
    df = df.dropna(subset=['Дата', 'Кол-во шт'], how='all')
    
    # Если нет региона, добавляем из других колонок
    if 'Регион' not in df.columns:
        # Пытаемся найти регион в других колонках
        for col in df.columns:
            if 'город' in str(col).lower() or 'область' in str(col).lower():
                df['Регион'] = df[col]
                break
    
    # Добавляем тип
    df['Тип'] = 'Биллинг'
    
    return df

def process_hired_data(df):
    """Обработка данных найма"""
    # Аналогичная обработка для найма
    for col in df.columns:
        col_lower = str(col).lower()
        if 'дата' in col_lower or 'date' in col_lower:
            df['Дата'] = df[col].apply(parse_date)
        elif 'шт' in col_lower or 'quantity' in col_lower or 'кол-во' in col_lower:
            df['Кол-во шт'] = pd.to_numeric(df[col], errors='coerce')
        elif 'регион' in col_lower or 'region' in col_lower:
            df['Регион'] = df[col]
        elif 'авто' in col_lower or 'тс' in col_lower or 'transport' in col_lower:
            df['Тип ТС'] = df[col]
        elif 'подрядчик' in col_lower:
            df['Подрядчик'] = df[col]
    
    df = df.dropna(subset=['Дата', 'Кол-во шт'], how='all')
    df['Тип'] = 'Наём'
    
    return df

def generate_test_data():
    """Генерация тестовых данных на основе вашей структуры"""
    np.random.seed(42)
    
    # Данные биллинга (из Общ таблица)
    dates = pd.date_range('2026-06-01', '2026-06-21', freq='D')
    regions = ['Ташкент', 'Фергана', 'Янгиколь', 'Сырдарья', 'Ташобласть']
    transports = ['газель', 'изсуз 5 тонн', 'Газель', 'Исусу 5т']
    stores = ['Магнит', 'АКО', 'Мафтуна', 'Базар Фергана', 'Мозгуново', 'ЦУМ Ташкент', 'Мега Планет']
    
    billing_data = []
    hired_data = []
    
    for date in dates:
        # Биллинг (данные от клиента)
        for _ in range(np.random.randint(2, 5)):
            region = np.random.choice(regions)
            transport = np.random.choice(transports)
            store = np.random.choice(stores)
            quantity = np.random.randint(1000, 20000)
            
            billing_data.append({
                'Дата': date,
                'Регион': region,
                'Тип ТС': transport,
                'Магазин': store,
                'Кол-во шт': quantity,
                '1 точка': np.random.randint(100000, 600000),
                '2 точка и далее': np.random.randint(0, 800000),
                'Грузчик экспедитор': np.random.randint(0, 500000),
                'Итого сумма': quantity * np.random.uniform(80, 150),
                'Тип': 'Биллинг'
            })
            
            # Наём (данные от подрядчиков)
            hired_data.append({
                'Дата': date,
                'Регион': region,
                'Тип ТС': transport,
                'Подрядчик': np.random.choice(['Premium Trade', 'Cargo Berg', 'Tadjiyev', 'Transaziya']),
                'Кол-во шт': int(quantity * np.random.uniform(0.7, 0.95)),
                'Кол-во точек': np.random.randint(1, 10),
                '1 точка': np.random.randint(100000, 500000),
                '2 точка и далее': np.random.randint(0, 700000),
                'Грузчик экспедиция': np.random.randint(0, 450000),
                'Итого сумма': quantity * np.random.uniform(50, 100),
                'Тип': 'Наём'
            })
    
    billing_df = pd.DataFrame(billing_data)
    hired_df = pd.DataFrame(hired_data)
    
    # Преобразуем даты
    billing_df['Дата'] = pd.to_datetime(billing_df['Дата'])
    hired_df['Дата'] = pd.to_datetime(hired_df['Дата'])
    
    return billing_df, hired_df

# Функция для создания Excel выгрузки
def create_excel_export(billing_df, hired_df, filename="analytics_export.xlsx"):
    """Создает Excel файл с данными"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Данные биллинга
        if not billing_df.empty:
            billing_export = billing_df.copy()
            if 'Дата' in billing_export.columns:
                billing_export['Дата'] = billing_export['Дата'].dt.strftime('%d.%m.%Y')
            billing_export.to_excel(writer, sheet_name='Биллинг', index=False)
        
        # Данные найма
        if not hired_df.empty:
            hired_export = hired_df.copy()
            if 'Дата' in hired_export.columns:
                hired_export['Дата'] = hired_export['Дата'].dt.strftime('%d.%m.%Y')
            hired_export.to_excel(writer, sheet_name='Наём', index=False)
            
            # Сравнение по дням
            if not billing_df.empty:
                daily_billing = billing_df.groupby('Дата')['Кол-во шт'].sum().reset_index()
                daily_billing.columns = ['Дата', 'Биллинг_шт']
                
                daily_hired = hired_df.groupby('Дата')['Кол-во шт'].sum().reset_index()
                daily_hired.columns = ['Дата', 'Наём_шт']
                
                comparison = pd.merge(daily_billing, daily_hired, on='Дата', how='outer').fillna(0)
                comparison['Разница'] = comparison['Биллинг_шт'] - comparison['Наём_шт']
                comparison['Дата'] = comparison['Дата'].dt.strftime('%d.%m.%Y')
                comparison.to_excel(writer, sheet_name='Сравнение_по_дням', index=False)
            
            # Сводка по подрядчикам
            if 'Подрядчик' in hired_df.columns:
                hired_by_contractor = hired_df.groupby('Подрядчик').agg({
                    'Кол-во шт': 'sum',
                    'Кол-во точек': 'sum' if 'Кол-во точек' in hired_df.columns else 'count'
                }).reset_index()
                hired_by_contractor.to_excel(writer, sheet_name='Наём_по_подрядчикам', index=False)
    
    output.seek(0)
    return output

# Основная функция
def main():
    # Заголовок
    st.markdown("""
        <div class="main-header">
            <h1>🚛 Аналитика логистики</h1>
            <p>📊 Доход от клиента (Биллинг) vs Расходы на наём (Подрядчики)</p>
        </div>
    """, unsafe_allow_html=True)
    
    # Загрузка данных
    with st.spinner("🔄 Загрузка данных..."):
        billing_df, hired_df = load_data_from_sheets()
    
    # Боковая панель
    with st.sidebar:
        st.markdown("""
            <div style="padding: 10px 0;">
                <h3 style="color: #00f5ff; margin-bottom: 20px;">🎯 Фильтры</h3>
            </div>
        """, unsafe_allow_html=True)
        
        # Выбор периода
        if not billing_df.empty and not hired_df.empty:
            all_dates = pd.concat([billing_df['Дата'], hired_df['Дата']])
            min_date = all_dates.min().date()
            max_date = all_dates.max().date()
        elif not billing_df.empty:
            min_date = billing_df['Дата'].min().date()
            max_date = billing_df['Дата'].max().date()
        else:
            min_date = datetime.now().date() - timedelta(days=30)
            max_date = datetime.now().date()
        
        date_range = st.date_input(
            "📅 Период",
            [min_date, max_date],
            min_value=min_date,
            max_value=max_date
        )
        
        if len(date_range) == 2:
            start_date, end_date = date_range
            if not billing_df.empty:
                billing_mask = (billing_df['Дата'].dt.date >= start_date) & (billing_df['Дата'].dt.date <= end_date)
                billing_filtered = billing_df.loc[billing_mask]
            else:
                billing_filtered = pd.DataFrame()
            
            if not hired_df.empty:
                hired_mask = (hired_df['Дата'].dt.date >= start_date) & (hired_df['Дата'].dt.date <= end_date)
                hired_filtered = hired_df.loc[hired_mask]
            else:
                hired_filtered = pd.DataFrame()
        else:
            billing_filtered = billing_df
            hired_filtered = hired_df
        
        # Фильтр по регионам
        all_regions = []
        if not billing_filtered.empty and 'Регион' in billing_filtered.columns:
            all_regions.extend(billing_filtered['Регион'].unique())
        if not hired_filtered.empty and 'Регион' in hired_filtered.columns:
            all_regions.extend(hired_filtered['Регион'].unique())
        
        all_regions = sorted(set(all_regions))
        
        if all_regions:
            regions = st.multiselect(
                "📍 Регионы",
                options=all_regions,
                default=all_regions
            )
            
            if regions:
                if not billing_filtered.empty and 'Регион' in billing_filtered.columns:
                    billing_filtered = billing_filtered[billing_filtered['Регион'].isin(regions)]
                if not hired_filtered.empty and 'Регион' in hired_filtered.columns:
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
    
    # Проверка наличия данных
    if billing_filtered.empty and hired_filtered.empty:
        st.warning("⚠️ Нет данных для отображения. Проверьте подключение к Google Sheets.")
        return
    
    # Ключевые метрики
    st.markdown("## 📊 Ключевые показатели")
    
    col1, col2, col3, col4 = st.columns(4)
    
    billing_total = billing_filtered['Кол-во шт'].sum() if not billing_filtered.empty else 0
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
                        <div class="metric-sub">{len(billing_filtered):,} записей</div>
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
                        <div class="metric-sub">{len(hired_filtered):,} записей</div>
                    </div>
                </div>
            </div>
        """, unsafe_allow_html=True)
    
    with col3:
        profit_class = "metric-value" if profit >= 0 else "metric-value-negative"
        profit_color = "profit-positive" if profit >= 0 else "profit-negative"
        st.markdown(f"""
            <div class="metric-card">
                <div style="display: flex; align-items: center; gap: 10px;">
                    <span style="font-size: 2rem;">📈</span>
                    <div>
                        <div class="{profit_class}">{profit:+,.0f}</div>
                        <div class="metric-label">Чистая прибыль (шт)</div>
                        <div class="metric-sub" style="color: {'#00ff88' if profit >= 0 else '#ff6b6b'}">
                            {profit/billing_total*100 if billing_total > 0 else 0:+.1f}% от биллинга
                        </div>
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
                        <div class="metric-sub">Биллинг / Наём</div>
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
        
        if not billing_filtered.empty:
            daily_billing = billing_filtered.groupby('Дата')['Кол-во шт'].sum().reset_index()
            daily_billing['Тип'] = 'Биллинг'
        else:
            daily_billing = pd.DataFrame()
        
        if not hired_filtered.empty:
            daily_hired = hired_filtered.groupby('Дата')['Кол-во шт'].sum().reset_index()
            daily_hired['Тип'] = 'Наём'
        else:
            daily_hired = pd.DataFrame()
        
        if not daily_billing.empty and not daily_hired.empty:
            daily_combined = pd.concat([daily_billing, daily_hired])
            
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
        else:
            st.info("Нет данных для графика динамики")
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.markdown("#### 🗺️ Распределение по регионам")
        
        if not billing_filtered.empty and 'Регион' in billing_filtered.columns:
            billing_regions = billing_filtered.groupby('Регион')['Кол-во шт'].sum().reset_index()
            billing_regions['Тип'] = 'Биллинг'
        else:
            billing_regions = pd.DataFrame()
        
        if not hired_filtered.empty and 'Регион' in hired_filtered.columns:
            hired_regions = hired_filtered.groupby('Регион')['Кол-во шт'].sum().reset_index()
            hired_regions['Тип'] = 'Наём'
        else:
            hired_regions = pd.DataFrame()
        
        if not billing_regions.empty and not hired_regions.empty:
            regions_combined = pd.concat([billing_regions, hired_regions])
            
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
        else:
            st.info("Нет данных для графика по регионам")
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Дополнительные графики
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.markdown("#### 🚚 Распределение по типам ТС")
        
        if not billing_filtered.empty and 'Тип ТС' in billing_filtered.columns:
            transport_data = billing_filtered.groupby('Тип ТС')['Кол-во шт'].sum().reset_index()
            
            fig = px.pie(
                transport_data,
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
        else:
            st.info("Нет данных по типам ТС")
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.markdown("#### 📊 Соотношение Биллинг vs Наём")
        
        if not billing_filtered.empty and not hired_filtered.empty:
            comparison_data = pd.DataFrame({
                'Категория': ['Биллинг', 'Наём'],
                'Кол-во шт': [billing_total, hired_total]
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
        else:
            st.info("Нет данных для сравнения")
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Детализация и выгрузка
    st.markdown("## 📋 Детализация и выгрузка")
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        if not billing_filtered.empty:
            detail_date = st.date_input(
                "📅 Выберите дату",
                value=billing_filtered['Дата'].max().date(),
                min_value=billing_filtered['Дата'].min().date(),
                max_value=billing_filtered['Дата'].max().date()
            )
        else:
            detail_date = datetime.now().date()
    
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("📥 Выгрузить в Excel", use_container_width=True):
            excel_file = create_excel_export(billing_filtered, hired_filtered)
            b64 = base64.b64encode(excel_file.getvalue()).decode()
            href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="analytics_export.xlsx" style="text-decoration: none;">📊 Скачать Excel</a>'
            st.markdown(href, unsafe_allow_html=True)
    
    # Детальная таблица
    if not billing_filtered.empty:
        detail_billing = billing_filtered[billing_filtered['Дата'].dt.date == detail_date]
        
        if not detail_billing.empty:
            # Выбираем колонки для отображения
            cols_to_show = ['Регион', 'Тип ТС', 'Кол-во шт']
            if 'Магазин' in detail_billing.columns:
                cols_to_show.append('Магазин')
            if '1 точка' in detail_billing.columns:
                cols_to_show.append('1 точка')
            if '2 точка и далее' in detail_billing.columns:
                cols_to_show.append('2 точка и далее')
            if 'Грузчик экспедитор' in detail_billing.columns:
                cols_to_show.append('Грузчик экспедитор')
            if 'Итого сумма' in detail_billing.columns:
                cols_to_show.append('Итого сумма')
            
            detail_table = detail_billing[cols_to_show].copy()
            
            # Добавляем данные по найму для этой даты
            if not hired_filtered.empty:
                detail_hired = hired_filtered[hired_filtered['Дата'].dt.date == detail_date]
                if not detail_hired.empty:
                    hired_summary = detail_hired.groupby('Регион')['Кол-во шт'].sum().reset_index()
                    hired_summary.columns = ['Регион', 'Наём_шт']
                    detail_table = detail_table.merge(hired_summary, on='Регион', how='left')
                    detail_table['Наём_шт'] = detail_table['Наём_шт'].fillna(0)
                    detail_table['Разница'] = detail_table['Кол-во шт'] - detail_table['Наём_шт']
            
            st.dataframe(
                detail_table,
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info("📭 Нет данных за выбранную дату")
    
    # Сводка по регионам
    st.markdown("## 📊 Сводка по регионам")
    
    if not billing_filtered.empty and 'Регион' in billing_filtered.columns:
        region_summary = billing_filtered.groupby('Регион').agg({
            'Кол-во шт': 'sum',
            'Магазин': 'nunique' if 'Магазин' in billing_filtered.columns else 'count'
        }).reset_index()
        region_summary.columns = ['Регион', 'Биллинг_шт', 'Кол-во_точек'] if len(region_summary.columns) == 3 else ['Регион', 'Биллинг_шт']
        
        if not hired_filtered.empty and 'Регион' in hired_filtered.columns:
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
            hide_index=True
        )
    
    # Информация о подрядчиках (если есть)
    if not hired_filtered.empty and 'Подрядчик' in hired_filtered.columns:
        st.markdown("## 👷 Статистика по подрядчикам")
        
        contractor_stats = hired_filtered.groupby('Подрядчик').agg({
            'Кол-во шт': 'sum',
            'Кол-во точек': 'sum' if 'Кол-во точек' in hired_filtered.columns else 'count'
        }).reset_index()
        
        st.dataframe(
            contractor_stats,
            use_container_width=True,
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
