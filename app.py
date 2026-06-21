import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import io
import base64
import json

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
    </style>
""", unsafe_allow_html=True)

# Загрузка данных из Google Sheets
@st.cache_data(ttl=3600)
def load_data():
    try:
        # Настройка доступа к Google Sheets
        scope = ['https://spreadsheets.google.com/feeds',
                'https://www.googleapis.com/auth/drive']
        
        # Используйте service account для доступа
        # ВАЖНО: Замените на ваши credentials
        creds_dict = {
            "type": "service_account",
            "project_id": "your-project-id",
            "private_key_id": "your-private-key-id",
            "private_key": "your-private-key",
            "client_email": "your-client-email",
            "client_id": "your-client-id",
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
            "client_x509_cert_url": "your-cert-url"
        }
        
        # Для демонстрации создаем тестовые данные
        # В реальном проекте используйте:
        # creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        # client = gspread.authorize(creds)
        # sheet = client.open_by_key('1-kv25KvN60XPksMm4YC27r_x9cxn0PevY4npWWdOi4c')
        
        # Создаем тестовые данные для демонстрации
        np.random.seed(42)
        dates = pd.date_range('2026-01-01', '2026-06-21', freq='D')
        
        data = []
        regions = ['Центральный', 'Северо-Западный', 'Южный', 'Приволжский', 'Уральский']
        transport_types = ['Грузовик', 'Фургон', 'Рефрижератор']
        stores = ['Магнит А', 'Магнит Б', 'Магнит В', 'Магнит Г', 'Магнит Д']
        
        for date in dates:
            for _ in range(np.random.randint(1, 5)):
                region = np.random.choice(regions)
                transport = np.random.choice(transport_types)
                store = np.random.choice(stores)
                quantity = np.random.randint(100, 500)
                amount = quantity * np.random.uniform(50, 150)
                
                data.append({
                    'Дата': date,
                    'Тип ТС': transport,
                    'Регион': region,
                    'Магазин': store,
                    'Кол-во шт': quantity,
                    'Сумма': round(amount, 2),
                    'Тип': np.random.choice(['Биллинг', 'Наём'])
                })
        
        df = pd.DataFrame(data)
        return df
    except Exception as e:
        st.error(f"Ошибка загрузки данных: {e}")
        return pd.DataFrame()

# Функция для создания выгрузки в Excel
def create_excel_download(df, filename="analytics_data.xlsx"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Данные', index=False)
        
        # Добавляем сводные таблицы
        pivot = df.pivot_table(
            values='Кол-во шт',
            index=['Дата', 'Регион'],
            columns='Тип ТС',
            aggfunc='sum',
            fill_value=0
        )
        pivot.to_excel(writer, sheet_name='Сводка по ТС')
        
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
    
    # Выбор периода
    min_date = df['Дата'].min().date()
    max_date = df['Дата'].max().date()
    
    date_range = st.sidebar.date_input(
        "Период",
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
    regions = st.sidebar.multiselect(
        "Регионы",
        options=sorted(filtered_df['Регион'].unique()),
        default=sorted(filtered_df['Регион'].unique())
    )
    
    if regions:
        filtered_df = filtered_df[filtered_df['Регион'].isin(regions)]
    
    transport = st.sidebar.multiselect(
        "Тип ТС",
        options=sorted(filtered_df['Тип ТС'].unique()),
        default=sorted(filtered_df['Тип ТС'].unique())
    )
    
    if transport:
        filtered_df = filtered_df[filtered_df['Тип ТС'].isin(transport)]
    
    # Кнопка обновления
    if st.sidebar.button("🔄 Обновить данные", type="primary"):
        st.cache_data.clear()
        st.rerun()
    
    # Основные метрики
    st.markdown("## 📊 Ключевые метрики")
    
    col1, col2, col3, col4 = st.columns(4)
    
    # Общие метрики
    total_quantity = filtered_df['Кол-во шт'].sum()
    total_amount = filtered_df['Сумма'].sum()
    avg_amount_per_unit = total_amount / total_quantity if total_quantity > 0 else 0
    
    # Метрики по типу
    billing_df = filtered_df[filtered_df['Тип'] == 'Биллинг']
    hired_df = filtered_df[filtered_df['Тип'] == 'Наём']
    
    billing_quantity = billing_df['Кол-во шт'].sum()
    hired_quantity = hired_df['Кол-во шт'].sum()
    billing_amount = billing_df['Сумма'].sum()
    hired_amount = hired_df['Сумма'].sum()
    
    with col1:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">Всего отгружено</div>
                <div class="metric-value">{total_quantity:,.0f} шт</div>
                <div style="color: rgba(255,255,255,0.5); font-size: 0.8rem;">на сумму {total_amount:,.0f} ₽</div>
            </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">Биллинг (клиент)</div>
                <div class="metric-value">{billing_quantity:,.0f} шт</div>
                <div style="color: #00f5ff; font-size: 0.9rem;">{billing_amount:,.0f} ₽</div>
                <div style="color: rgba(255,255,255,0.5); font-size: 0.8rem;">средняя цена за шт: {billing_amount/billing_quantity if billing_quantity > 0 else 0:,.2f} ₽</div>
            </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">Наём (исполнители)</div>
                <div class="metric-value">{hired_quantity:,.0f} шт</div>
                <div style="color: #ff6b6b; font-size: 0.9rem;">{hired_amount:,.0f} ₽</div>
                <div style="color: rgba(255,255,255,0.5); font-size: 0.8rem;">средняя цена за шт: {hired_amount/hired_quantity if hired_quantity > 0 else 0:,.2f} ₽</div>
            </div>
        """, unsafe_allow_html=True)
    
    with col4:
        profit = billing_amount - hired_amount
        profit_per_unit = profit / total_quantity if total_quantity > 0 else 0
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">Маржинальная прибыль</div>
                <div class="metric-value" style="color: {'#00ff88' if profit > 0 else '#ff6b6b'}">
                    {profit:,.0f} ₽
                </div>
                <div style="color: rgba(255,255,255,0.5); font-size: 0.8rem;">
                    прибыль за шт: {profit_per_unit:,.2f} ₽
                </div>
                <div style="color: rgba(255,255,255,0.5); font-size: 0.8rem;">
                    маржинальность: {(profit/billing_amount*100) if billing_amount > 0 else 0:.1f}%
                </div>
            </div>
        """, unsafe_allow_html=True)
    
    # Графики в 2 колонки
    st.markdown("## 📈 Аналитические графики")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # График сравнения Биллинг vs Наём
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
            textfont=dict(color='white')
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
            height=350
        )
        st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        # Распределение по регионам
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.subheader("Распределение по регионам")
        
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
        fig.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white',
            showlegend=True,
            legend=dict(
                orientation="v",
                yanchor="middle",
                y=0.5,
                xanchor="left",
                x=1.1
            ),
            margin=dict(l=0, r=100, t=30, b=0),
            height=350
        )
        st.plotly_chart(fig, use_container_width=True)
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
            marker=dict(size=6, color='#00f5ff')
        ))
        fig.add_trace(go.Scatter(
            x=daily_stats['Дата'],
            y=daily_stats['Сумма'],
            mode='lines+markers',
            name='Сумма ₽',
            yaxis='y2',
            line=dict(color='#ffd93d', width=3),
            marker=dict(size=6, color='#ffd93d')
        ))
        
        fig.update_layout(
            yaxis=dict(title="Кол-во шт", titlefont=dict(color='#00f5ff'), tickfont=dict(color='#00f5ff'), gridcolor='rgba(255,255,255,0.1)'),
            yaxis2=dict(title="Сумма ₽", titlefont=dict(color='#ffd93d'), tickfont=dict(color='#ffd93d'), overlaying='y', side='right', gridcolor='rgba(255,255,255,0.1)'),
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
            height=350
        )
        st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.subheader("Распределение по типам ТС")
        
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
            color_discrete_sequence=px.colors.sequential.Plasma_r
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
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Анализ эффективности
    st.markdown("## 💰 Анализ эффективности")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.subheader("Средняя цена за шт")
        
        billing_avg = billing_amount / billing_quantity if billing_quantity > 0 else 0
        hired_avg = hired_amount / hired_quantity if hired_quantity > 0 else 0
        
        fig = go.Figure()
        fig.add_trace(go.Indicator(
            mode="gauge+number+delta",
            value=billing_avg,
            delta={'reference': hired_avg, 'increasing': {'color': "#00ff88"}, 'decreasing': {'color': "#ff6b6b"}},
            gauge={
                'axis': {'range': [None, max(billing_avg, hired_avg) * 1.2], 'tickfont': {'color': 'white'}},
                'bar': {'color': "#00f5ff"},
                'steps': [
                    {'range': [0, hired_avg], 'color': 'rgba(255,107,107,0.3)'},
                    {'range': [hired_avg, max(billing_avg, hired_avg) * 1.2], 'color': 'rgba(0,245,255,0.3)'}
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
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.subheader("Маржинальность")
        
        margin = ((billing_amount - hired_amount) / billing_amount * 100) if billing_amount > 0 else 0
        
        fig = go.Figure(go.Indicator(
            mode = "gauge+number",
            value = margin,
            title = {'text': "Маржинальность %", 'font': {'color': 'white', 'size': 16}},
            number = {'suffix': "%", 'font': {'color': 'white', 'size': 40}},
            gauge = {
                'axis': {'range': [0, 100], 'tickfont': {'color': 'white'}},
                'bar': {'color': "#00ff88" if margin > 20 else "#ffd93d" if margin > 10 else "#ff6b6b"},
                'steps': [
                    {'range': [0, 30], 'color': 'rgba(255,107,107,0.2)'},
                    {'range': [30, 70], 'color': 'rgba(255,217,61,0.2)'},
                    {'range': [70, 100], 'color': 'rgba(0,255,136,0.2)'}
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
        fig.update_traces(texttemplate='%{text:.3f}', textposition='outside', textfont_color='white')
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
            "Выберите дату для детализации",
            value=filtered_df['Дата'].max().date(),
            min_value=filtered_df['Дата'].min().date(),
            max_value=filtered_df['Дата'].max().date()
        )
    
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("📥 Выгрузить все данные в Excel", type="primary"):
            excel_file = create_excel_download(filtered_df)
            b64 = base64.b64encode(excel_file.getvalue()).decode()
            href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="analytics_data.xlsx">Скачать Excel</a>'
            st.markdown(href, unsafe_allow_html=True)
    
    # Детальная таблица
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
                'Регион': 'Регион',
                'Тип ТС': 'Тип ТС',
                'Тип': 'Тип',
                'Магазин': 'Магазины',
                'Кол-во шт': st.column_config.NumberColumn('Кол-во шт', format='%d'),
                'Сумма': st.column_config.NumberColumn('Сумма ₽', format='% .2f')
            },
            hide_index=True
        )
    else:
        st.info("Нет данных за выбранную дату")
    
    # Дополнительная информация
    with st.expander("📌 Дополнительная информация"):
        st.markdown("""
            **Аналитические выводы:**
            - 💡 Сравнительный анализ показывает разницу между доходами от клиентов и расходами на наём
            - 📊 Графики демонстрируют распределение по регионам и типам транспортных средств
            - 🎯 Маржинальность показывает эффективность бизнес-модели
            - 📈 Динамика позволяет отслеживать тренды и сезонность
            
            **Структура данных:**
            - Биллинг - доходы от клиента (Магнит)
            - Наём - расходы на исполнителей
            - Маржинальная прибыль = Биллинг - Наём
        """)

if __name__ == "__main__":
    main()
