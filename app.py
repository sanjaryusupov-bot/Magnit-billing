import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

from data_parsing import (
    add_route_economics,
    attach_city,
    build_registry_export,
    build_sla,
    build_sla_export,
    load_from_gsheet_service_account,
)

st.set_page_config(page_title="MCOS — статус по отгрузкам с РЦ", layout="wide", page_icon="📦")

# Единая цветовая палитра для графиков
COLOR_QTY = "#4C78A8"
COLOR_SUM = "#F58518"
COLOR_BG = "rgba(0,0,0,0)"

CUSTOM_CSS = """
<style>
div[data-testid="stMetric"] {
    background: rgba(76, 120, 168, 0.08);
    border: 1px solid rgba(76, 120, 168, 0.18);
    border-radius: 10px;
    padding: 12px 16px 8px 16px;
}
div[data-testid="stMetricValue"] { font-size: 1.6rem; }
.stTabs [data-baseweb="tab-list"] { gap: 6px; }
.stTabs [data-baseweb="tab"] {
    border-radius: 8px 8px 0 0;
    padding: 8px 16px;
}
div[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def fmt_int(v):
    try:
        return f"{int(round(v)):,}".replace(",", " ")
    except (TypeError, ValueError):
        return "—"


def fmt_money(v):
    try:
        return f"{v:,.0f}".replace(",", " ")
    except (TypeError, ValueError):
        return "—"


def fmt_rate(v):
    try:
        return f"{v:,.2f}".replace(",", " ")
    except (TypeError, ValueError):
        return "—"


@st.cache_data(show_spinner="Загружаю Google Sheet...", ttl=300)
def _load(sheet_id: str, _creds_marker: str):
    creds_info = dict(st.secrets["gcp_service_account"])
    return load_from_gsheet_service_account(sheet_id, creds_info)


# ---------------------------------------------------------------------------
# Sidebar — источник данных: Google Sheet через сервисный аккаунт (ID таблицы
# зашит в код — поле для его ввода в интерфейсе не показывается)
# ---------------------------------------------------------------------------
SHEET_ID = "12CxJMCBUHgkaj-_KbOs1aK7hx_jEYMyS3q4Hh0bHTGw"

st.sidebar.header("📦 MCOS")

ship_df = pd.DataFrame()
orders_df = pd.DataFrame()
tariffs_df = pd.DataFrame()

if st.sidebar.button("🔄 Загрузить / обновить данные", use_container_width=True, type="primary"):
    _load.clear()

if "gcp_service_account" not in st.secrets:
    st.sidebar.error(
        "Не настроены Secrets приложения (gcp_service_account). "
        "Добавьте их в Manage app → Settings → Secrets."
    )
else:
    try:
        client_email = st.secrets["gcp_service_account"].get("client_email", "")
        ship_df, orders_df, tariffs_df = _load(SHEET_ID, client_email)
        if tariffs_df.empty:
            st.sidebar.info(
                "В таблице нет листа «тарифы» — в экспорте будет использован "
                "встроенный справочник ставок по умолчанию."
            )
    except Exception as e:
        msg = str(e)
        if "PERMISSION_DENIED" in msg or "403" in msg:
            st.sidebar.error(
                f"Нет доступа к таблице. Откройте доступ на чтение для {client_email or 'сервисного аккаунта'} "
                "(Настройки доступа таблицы → Читатель)."
            )
        elif "WorksheetNotFound" in msg or "worksheet" in msg.lower():
            st.sidebar.error(f"Не найден один из листов ('Статус по отгрузкам' / 'Статус по заказам'): {e}")
        else:
            st.sidebar.error(f"Не удалось загрузить таблицу: {e}")

if ship_df.empty:
    st.title("📦 MCOS — статус по отгрузкам с РЦ")
    st.info("⬅️ Слева нажмите «Загрузить / обновить данные», чтобы увидеть дашборд.")
    st.stop()

# Экономика маршрута считается на ПОЛНЫХ (нефильтрованных) данных: тариф за
# штуку должен учитывать все заказы рейса, а не только те, что останутся
# после фильтра периода. Также только заказы 1M-... (всегда, без галочки —
# остальные в отчётах не нужны).
ship_df = ship_df[ship_df["Заказ"].str.startswith("1M-")].copy()
ship_df = add_route_economics(ship_df)
ship_df = attach_city(ship_df, orders_df)

# ---------------------------------------------------------------------------
# Глобальные фильтры
# ---------------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.header("🔎 Фильтры")

min_d, max_d = ship_df["Дата"].min().date(), ship_df["Дата"].max().date()
date_range = st.sidebar.date_input("Период (по дате отгрузки)", value=(min_d, max_d), min_value=min_d, max_value=max_d)
if isinstance(date_range, tuple) and len(date_range) == 2:
    start_d, end_d = date_range
else:
    start_d, end_d = min_d, max_d

mask = (ship_df["Дата"].dt.date >= start_d) & (ship_df["Дата"].dt.date <= end_d)
f = ship_df[mask].copy()

st.title("📦 MCOS — статус по отгрузкам с РЦ")
st.caption(
    "Стоимость доставки распределяется по рейсу пропорционально количеству штук: "
    "Тариф за шт = Сумма тарифов по всем точкам рейса / Кол-во штук в рейсе. "
    "Так сумма по каждому магазину/заказу считается честно, даже если заказов "
    "в один магазин за день несколько."
)

tab1, tab2, tab3, tab4 = st.tabs(
    [
        "📅 Отгрузки по дням",
        "🗺️ Сводная (Город / Транспорт / Рейс)",
        "⏱️ SLA: план vs факт",
        "📤 Экспорт в Excel",
    ]
)

# ---------------------------------------------------------------------------
# TAB 1 — Отгрузки по дням + переход в детали
# ---------------------------------------------------------------------------
with tab1:
    st.subheader("Отгрузки по дням")

    daily = (
        f.groupby("Дата", as_index=False)
        .agg(Кол_во_шт=("Кол_во_шт", "sum"), Итого_сумма=("Сумма_распределенная", "sum"))
        .sort_values("Дата")
    )

    c1, c2, c3 = st.columns(3)
    c1.metric("Дней в периоде", len(daily))
    c2.metric("Итого штук", fmt_int(daily["Кол_во_шт"].sum()))
    c3.metric("Итого сумма", fmt_money(daily["Итого_сумма"].sum()))

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(
        go.Bar(
            x=daily["Дата"], y=daily["Кол_во_шт"], name="Кол-во шт",
            marker_color=COLOR_QTY, opacity=0.85,
            hovertemplate="%{x|%d.%m.%Y}<br>Кол-во шт: %{y:,.0f}<extra></extra>",
        ),
        secondary_y=False,
    )
    fig.add_trace(
        go.Scatter(
            x=daily["Дата"], y=daily["Итого_сумма"], name="Итого сумма",
            mode="lines+markers", line=dict(color=COLOR_SUM, width=3), marker=dict(size=6),
            hovertemplate="%{x|%d.%m.%Y}<br>Сумма: %{y:,.0f}<extra></extra>",
        ),
        secondary_y=True,
    )
    fig.update_layout(
        height=420,
        margin=dict(l=10, r=10, t=40, b=10),
        plot_bgcolor=COLOR_BG,
        paper_bgcolor=COLOR_BG,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        hovermode="x unified",
        title=dict(text="Кол-во штук и сумма по дням", x=0.01, xanchor="left"),
    )
    fig.update_yaxes(title_text="Кол-во шт", secondary_y=False, gridcolor="rgba(128,128,128,0.15)")
    fig.update_yaxes(title_text="Сумма", secondary_y=True, showgrid=False)
    fig.update_xaxes(gridcolor="rgba(128,128,128,0.10)")
    st.plotly_chart(fig, use_container_width=True)

    st.dataframe(
        daily.rename(columns={"Дата": "День"}),
        use_container_width=True,
        hide_index=True,
        column_config={
            "День": st.column_config.DateColumn(format="DD.MM.YYYY"),
            "Кол_во_шт": st.column_config.NumberColumn("Кол-во шт", format="%d"),
            "Итого_сумма": st.column_config.NumberColumn("Итого сумма", format="%.0f"),
        },
    )

    # Детали по магазинам за ВЕСЬ период (используется и для drill-down, и для экспорта)
    store_detail_all = (
        f.groupby(["Дата", "Магазин", "Город"], as_index=False, dropna=False)
        .agg(Кол_во_шт=("Кол_во_шт", "sum"), Сумма=("Сумма_распределенная", "sum"))
    )
    store_detail_all["Сумма_за_шт"] = store_detail_all.apply(
        lambda r: (r["Сумма"] / r["Кол_во_шт"]) if r["Кол_во_шт"] else 0.0, axis=1
    )

    st.markdown("#### 🔽 Переход в детали по дню")
    if not daily.empty:
        day_options = daily["Дата"].dt.date.tolist()
        sel_day = st.selectbox("Выберите день", day_options, format_func=lambda d: d.strftime("%d.%m.%Y"))
        detail = (
            store_detail_all[store_detail_all["Дата"].dt.date == sel_day]
            .drop(columns=["Дата"])
            .assign(Дата=sel_day)
        )
        detail = detail[["Дата", "Магазин", "Город", "Кол_во_шт", "Сумма", "Сумма_за_шт"]].sort_values(
            "Сумма", ascending=False
        )
        st.dataframe(
            detail,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Дата": st.column_config.DateColumn(format="DD.MM.YYYY"),
                "Кол_во_шт": st.column_config.NumberColumn("Кол-во шт", format="%d"),
                "Сумма": st.column_config.NumberColumn(format="%.0f"),
                "Сумма_за_шт": st.column_config.NumberColumn("Сумма за шт", format="%.2f"),
            },
        )
        st.download_button(
            "⬇️ Скачать детали дня (CSV)",
            detail.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"detali_{sel_day}.csv",
            mime="text/csv",
        )

# ---------------------------------------------------------------------------
# TAB 2 — Сводная: Город / Транспорт / Рейс / Дата
# ---------------------------------------------------------------------------
with tab2:
    st.subheader("Сводная: Город / Транспорт / Рейс")

    merged = f.copy()

    fc1, fc2 = st.columns(2)
    with fc1:
        city_opts = sorted([c for c in merged["Город"].dropna().unique()])
        sel_cities = st.multiselect("Город", city_opts)
    with fc2:
        transp_opts = sorted([t for t in merged["Транспорт"].dropna().unique()])
        sel_transp = st.multiselect("Транспорт", transp_opts)

    if sel_cities:
        merged = merged[merged["Город"].isin(sel_cities)]
    if sel_transp:
        merged = merged[merged["Транспорт"].isin(sel_transp)]

    summary = (
        merged.groupby(["Город", "Транспорт", "Маршрут_ID", "Дата"], dropna=False, as_index=False)
        .agg(Итого_сумма=("Сумма_распределенная", "sum"), Кол_во_шт=("Кол_во_шт", "sum"))
        .rename(columns={"Маршрут_ID": "Рейс"})
        .sort_values("Итого_сумма", ascending=False)
    )
    summary["Сумма_за_шт"] = summary.apply(
        lambda r: (r["Итого_сумма"] / r["Кол_во_шт"]) if r["Кол_во_шт"] else 0.0, axis=1
    )
    summary = summary[["Дата", "Город", "Транспорт", "Рейс", "Итого_сумма", "Кол_во_шт", "Сумма_за_шт"]]

    c1, c2 = st.columns(2)
    c1.metric("Итого штук", fmt_int(summary["Кол_во_шт"].sum()))
    c2.metric("Итого сумма", fmt_money(summary["Итого_сумма"].sum()))

    st.dataframe(
        summary,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Дата": st.column_config.DateColumn(format="DD.MM.YYYY"),
            "Кол_во_шт": st.column_config.NumberColumn("Кол-во шт", format="%d"),
            "Итого_сумма": st.column_config.NumberColumn("Итого сумма", format="%.0f"),
            "Сумма_за_шт": st.column_config.NumberColumn("Сумма за шт", format="%.2f"),
        },
    )
    st.download_button(
        "⬇️ Скачать сводную (CSV)",
        summary.to_csv(index=False).encode("utf-8-sig"),
        file_name="svodnaya_gorod_transport_reys.csv",
        mime="text/csv",
    )

    st.caption(
        "Рейс — это Маршрут_ID (номер первого заказа маршрута), уникальный по всей выгрузке. "
        "Город берётся из листа «Статус по заказам» по номеру заказа; если заказ там не найден — "
        "используется колонка «Регион» из «Статус по отгрузкам». Сумма — распределённая "
        "пропорционально кол-ву штук (см. пояснение вверху страницы)."
    )

# ---------------------------------------------------------------------------
# TAB 3 — SLA: план vs факт
# ---------------------------------------------------------------------------
with tab3:
    st.subheader("SLA: плановая дата отгрузки vs фактическая")
    st.caption(
        "План — из листа «Статус по заказам» (колонка «Дата отгрузки план», "
        "либо дата раздела, если колонка не заполнена). "
        "Факт — дата из листа «Статус по отгрузкам». Сравнение только по заказам 1M-..."
    )

    sla = pd.DataFrame()
    if orders_df.empty:
        st.warning("Нет данных листа «Статус по заказам» — SLA посчитать не из чего.")
    else:
        sla = build_sla(ship_df, orders_df)

        status_opts = sla["SLA_статус"].unique().tolist()
        sel_status = st.multiselect("Статус SLA", status_opts, default=status_opts)
        sla_f = sla[sla["SLA_статус"].isin(sel_status)] if sel_status else sla

        c1, c2, c3, c4 = st.columns(4)
        total = len(sla_f)
        on_time = (sla_f["SLA_статус"] == "В срок").sum()
        late = (sla_f["SLA_статус"] == "Просрочка").sum()
        not_shipped = (sla_f["SLA_статус"] == "Не отгружен").sum()
        c1.metric("Всего заказов", total)
        c2.metric("В срок", on_time, f"{(on_time/total*100):.0f}%" if total else "0%")
        c3.metric("Просрочка", late, f"{(late/total*100):.0f}%" if total else "0%")
        c4.metric("Не отгружен", not_shipped, f"{(not_shipped/total*100):.0f}%" if total else "0%")

        status_counts = sla_f["SLA_статус"].value_counts()
        status_colors = {"В срок": "#54A24B", "Просрочка": "#E45756", "Не отгружен": "#B0B0B0", "Нет плана": "#EECA3B"}
        fig3 = go.Figure(
            go.Pie(
                labels=status_counts.index,
                values=status_counts.values,
                hole=0.55,
                marker=dict(colors=[status_colors.get(s, "#4C78A8") for s in status_counts.index]),
                textinfo="label+percent",
            )
        )
        fig3.update_layout(
            height=360,
            margin=dict(l=10, r=10, t=30, b=10),
            paper_bgcolor=COLOR_BG,
            showlegend=False,
        )
        st.plotly_chart(fig3, use_container_width=True)

        st.dataframe(
            sla_f.sort_values("Дельта_дней", ascending=False),
            use_container_width=True,
            hide_index=True,
            column_config={
                "Дата_план": st.column_config.DateColumn("Дата план", format="DD.MM.YYYY"),
                "Дата_факт": st.column_config.DateColumn("Дата факт", format="DD.MM.YYYY"),
                "Дельта_дней": st.column_config.NumberColumn("Дельта, дн."),
            },
        )

        dc1, dc2 = st.columns(2)
        with dc1:
            st.download_button(
                "⬇️ Скачать SLA (CSV)",
                sla_f.to_csv(index=False).encode("utf-8-sig"),
                file_name="sla_plan_vs_fakt.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with dc2:
            st.download_button(
                "📥 Скачать SLA (Excel)",
                data=build_sla_export(sla_f),
                file_name=f"sla_{start_d}_{end_d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

# ---------------------------------------------------------------------------
# TAB 4 — Экспорт реестра в Excel
# ---------------------------------------------------------------------------
with tab4:
    st.subheader("Экспорт в Excel")
    st.markdown(
        """
Один лист **«Реестр»** — построчно по каждому заказу за выбранный период
(оформлен как таблица с фильтрами, без колонки «Подрядчик»):

Дата отгрузки, Заказ, Рейс (Маршрут_ID — номер первого заказа маршрута,
уникален по всей выгрузке), Кол-во штук, ID магазина, Магазин, Адрес,
Координаты, Регион, Регион тариф, Город, Авто, 1 точка, 2 точка и далее,
Грузчик экспедитор, Тариф (исходный), Итого сумма (честно распределённая
по кол-ву штук — см. пояснение вверху страницы).

SLA (план vs факт) выгружается отдельным файлом — на вкладке
«⏱️ SLA: план vs факт».
        """
    )

    registry_bytes = build_registry_export(f)
    st.download_button(
        "📥 Скачать реестр (Excel)",
        data=registry_bytes,
        file_name=f"reestr_{start_d}_{end_d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

    registry_preview = f.sort_values(["Дата", "Маршрут_ID", "Заказ"])[
        ["Дата", "Заказ", "Маршрут_ID", "Кол_во_шт", "Магазин", "Город", "Регион", "Регион_тариф",
         "Транспорт", "Тариф_за_шт", "Сумма_распределенная"]
    ].rename(columns={"Маршрут_ID": "Рейс", "Сумма_распределенная": "Итого_сумма"})

    st.markdown("##### Предпросмотр реестра (первые 50 строк)")
    st.dataframe(
        registry_preview.head(50),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Дата": st.column_config.DateColumn(format="DD.MM.YYYY"),
            "Кол_во_шт": st.column_config.NumberColumn("Кол-во шт", format="%d"),
            "Тариф_за_шт": st.column_config.NumberColumn(format="%.2f"),
            "Итого_сумма": st.column_config.NumberColumn(format="%.0f"),
        },
    )
