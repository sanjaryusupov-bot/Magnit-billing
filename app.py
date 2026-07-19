import pandas as pd
import streamlit as st

from data_parsing import load_from_gsheet_service_account, build_sla

st.set_page_config(page_title="MCOS — статус по отгрузкам с РЦ", layout="wide", page_icon="📦")


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


@st.cache_data(show_spinner="Загружаю Google Sheet...", ttl=300)
def _load(sheet_id: str, _creds_marker: str):
    # _creds_marker (client_email) участвует в кэш-ключе, реальные креды берём из secrets внутри
    creds_info = dict(st.secrets["gcp_service_account"])
    return load_from_gsheet_service_account(sheet_id, creds_info)


# ---------------------------------------------------------------------------
# Sidebar — источник данных: Google Sheet через сервисный аккаунт
# ---------------------------------------------------------------------------
st.sidebar.header("⚙️ Google Sheet")
default_id = "12CxJMCBUHgkaj-_KbOs1aK7hx_jEYMyS3q4Hh0bHTGw"
sheet_id = st.sidebar.text_input(
    "ID Google-таблицы",
    value=default_id,
    help="ID — часть ссылки между /d/ и /edit.",
)

ship_df = pd.DataFrame()
orders_df = pd.DataFrame()

if st.sidebar.button("🔄 Загрузить / обновить данные", use_container_width=True):
    _load.clear()

if "gcp_service_account" not in st.secrets:
    st.sidebar.error(
        "Не настроены Secrets приложения (gcp_service_account). "
        "Добавьте их в Manage app → Settings → Secrets."
    )
elif sheet_id:
    try:
        client_email = st.secrets["gcp_service_account"].get("client_email", "")
        ship_df, orders_df = _load(sheet_id, client_email)
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
    st.info("⬅️ Слева загрузите Excel-файл или подтяните данные из Google Sheet, чтобы увидеть дашборд.")
    st.stop()

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

only_1m = st.sidebar.checkbox("Только заказы 1M-...", value=True)

mask = (ship_df["Дата"].dt.date >= start_d) & (ship_df["Дата"].dt.date <= end_d)
if only_1m:
    mask &= ship_df["Заказ"].str.startswith("1M-")
f = ship_df[mask].copy()

st.title("📦 MCOS — статус по отгрузкам с РЦ")

tab1, tab2, tab3 = st.tabs(
    ["📅 Отгрузки по дням", "🗺️ Сводная (Город / Транспорт / Рейс)", "⏱️ SLA: план vs факт"]
)

# ---------------------------------------------------------------------------
# TAB 1 — Отгрузки по дням + переход в детали
# ---------------------------------------------------------------------------
with tab1:
    st.subheader("Отгрузки по дням")

    daily = (
        f.groupby("Дата", as_index=False)
        .agg(Кол_во_шт=("Кол_во_шт", "sum"), Итого_сумма=("Итого_сумма", "sum"))
        .sort_values("Дата")
    )

    c1, c2, c3 = st.columns(3)
    c1.metric("Дней в периоде", len(daily))
    c2.metric("Итого штук", fmt_int(daily["Кол_во_шт"].sum()))
    c3.metric("Итого сумма", fmt_money(daily["Итого_сумма"].sum()))

    cc1, cc2 = st.columns(2)
    with cc1:
        st.caption("Кол-во штук по дням")
        st.bar_chart(daily.set_index("Дата")[["Кол_во_шт"]])
    with cc2:
        st.caption("Итого сумма по дням")
        st.bar_chart(daily.set_index("Дата")[["Итого_сумма"]])

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

    st.markdown("#### 🔽 Переход в детали по дню")
    if not daily.empty:
        day_options = daily["Дата"].dt.date.tolist()
        sel_day = st.selectbox("Выберите день", day_options, format_func=lambda d: d.strftime("%d.%m.%Y"))
        detail = (
            f[f["Дата"].dt.date == sel_day]
            .groupby("Магазин", as_index=False, dropna=False)
            .agg(Кол_во_шт=("Кол_во_шт", "sum"), Сумма=("Итого_сумма", "sum"))
            .sort_values("Сумма", ascending=False)
        )
        detail.insert(0, "Дата", sel_day)
        st.dataframe(
            detail,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Дата": st.column_config.DateColumn(format="DD.MM.YYYY"),
                "Кол_во_шт": st.column_config.NumberColumn("Кол-во шт", format="%d"),
                "Сумма": st.column_config.NumberColumn(format="%.0f"),
            },
        )
        st.download_button(
            "⬇️ Скачать детали дня (CSV)",
            detail.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"detali_{sel_day}.csv",
            mime="text/csv",
        )

# ---------------------------------------------------------------------------
# TAB 2 — Сводная: Город / Транспорт / Рейс
# ---------------------------------------------------------------------------
with tab2:
    st.subheader("Сводная: Город / Транспорт / Рейс")

    merged = f.copy()
    if not orders_df.empty:
        merged = merged.merge(orders_df[["Заказ", "Город"]], on="Заказ", how="left")
    else:
        merged["Город"] = None
    merged["Город_итог"] = merged["Город"].fillna(merged["Регион"])

    fc1, fc2 = st.columns(2)
    with fc1:
        city_opts = sorted([c for c in merged["Город_итог"].dropna().unique()])
        sel_cities = st.multiselect("Город", city_opts)
    with fc2:
        transp_opts = sorted([t for t in merged["Транспорт"].dropna().unique()])
        sel_transp = st.multiselect("Транспорт", transp_opts)

    if sel_cities:
        merged = merged[merged["Город_итог"].isin(sel_cities)]
    if sel_transp:
        merged = merged[merged["Транспорт"].isin(sel_transp)]

    summary = (
        merged.groupby(["Город_итог", "Транспорт", "Рейс"], dropna=False, as_index=False)
        .agg(Итого_сумма=("Итого_сумма", "sum"), Кол_во_шт=("Кол_во_шт", "sum"))
        .rename(columns={"Город_итог": "Город"})
        .sort_values("Итого_сумма", ascending=False)
    )

    c1, c2 = st.columns(2)
    c1.metric("Итого штук", fmt_int(summary["Кол_во_шт"].sum()))
    c2.metric("Итого сумма", fmt_money(summary["Итого_сумма"].sum()))

    st.dataframe(
        summary,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Кол_во_шт": st.column_config.NumberColumn("Кол-во шт", format="%d"),
            "Итого_сумма": st.column_config.NumberColumn("Итого сумма", format="%.0f"),
        },
    )
    st.download_button(
        "⬇️ Скачать сводную (CSV)",
        summary.to_csv(index=False).encode("utf-8-sig"),
        file_name="svodnaya_gorod_transport_reys.csv",
        mime="text/csv",
    )

    st.caption(
        "Город берётся из листа «Статус по заказам» по номеру заказа; если заказ там не найден — "
        "используется колонка «Регион» из «Статус по отгрузкам»."
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
        st.download_button(
            "⬇️ Скачать SLA (CSV)",
            sla_f.to_csv(index=False).encode("utf-8-sig"),
            file_name="sla_plan_vs_fakt.csv",
            mime="text/csv",
        )
