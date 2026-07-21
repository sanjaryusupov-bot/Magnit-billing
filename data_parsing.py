"""
Парсинг листов гугл-таблицы MCOS в плоские (tidy) таблицы для дашборда.

Лист "Статус по отгрузкам" — это визуальный отчёт с группировками:
    День (строка с датой, без заказа)
      Рейс N (строка-заголовок)
        строки заказов (дата, заказ, кол-во шт, магазин, регион, авто, сумма...)

Лист "Статус по заказам" — тоже с группировками по дням в колонке A:
    "01/07/2026 (заказы на ср)"  <- заголовок раздела (это и есть план. дата)
      строки заказов (№ заказа, ID магазина, магазин, кол-во шт, дата отгрузки план, город, статусы...)
"""
import re
from datetime import datetime, date, timedelta
from io import BytesIO

import openpyxl
import pandas as pd

SHIPMENTS_SHEET = "Статус по отгрузкам"
ORDERS_SHEET = "Статус по заказам"
TARIFFS_SHEET = "тарифы"

_DATE_RE = re.compile(r"(\d{2})/(\d{2})/(\d{4})")

# Колонки (1-based), которые в исходных листах содержат дату и должны быть
# приведены к datetime, даже когда данные пришли как "сырые" числа (Google
# Sheets API отдаёт даты как порядковый номер дня, как в Excel).
_SHIPMENTS_DATE_COLS = {2}   # B: Дата отгрузки
_ORDERS_DATE_COLS = {6}      # F: Дата отгрузки план

_EXCEL_EPOCH = datetime(1899, 12, 30)

# Регион (точная точка доставки) -> Регион тариф (укрупнённая тарифная зона).
# По умолчанию — маппинг, извлечённый из реального файла биллинга заказчика
# (спутниковые города/районы тарифицируются по ближайшему хабу). Значения,
# которых нет в словаре, считаются тарифной зоной "как есть" (сами собой).
DEFAULT_REGION_TARIFF_MAP = {
    "Ангрен": "Ташкент",
    "Сырдарья": "Ташкент",
    "Ташобласть - Нурафшон": "Ташкент",
    "Ташобласть - Чирчик": "Ташкент",
    "Янгиюль": "Ташкент",
    "Коканд": "Фергана",
}

# Справочник тарифов по умолчанию (копия листа "тарифы" из примера биллинга
# заказчика). Используется, если в самой Google-таблице нет листа "тарифы" —
# тогда экспорт всё равно будет содержать актуальный на момент настройки
# справочник. Если тарифы поменяются, обновите эту таблицу или добавьте лист
# "тарифы" в саму Google-таблицу — он будет использован автоматически.
DEFAULT_TARIFFS = [
    (1, "Ташкент", "Лабо с грузоподъемностью до 500кг. (9:00 - 18:00)",
     "Доставка заказов по г. Ташкент овтомобилем Лабо с грузоподъемностью до 500кг. (9:00 - 18:00)", 400000, 55000),
    (2, "Ташкент", "Лабо", "Доставка заказов по городу Ташкент - Лабо", 210000, 55000),
    (3, "Ташкент", "Газель", "Доставка заказов по городу Ташкент - Газель", 450000, 165000),
    (4, "Ташкент", "Исузу 5 тонн", "Доставка заказов по городу Ташкент - Исузу 5тонн", 935000, 165000),
    (5, "Ташкент", "10 тонн", "Доставка заказов по городу - 10 тонн", 1430000, 165000),
    (6, "Ташкент", "Евро фура", "Доставка заказов по городу - Евро фура", 1760000, 220000),
    (7, "Самарканд", "Газель", "Ташкент-Самарканд - Газель", 1900000, 165000),
    (8, "Самарканд", "Исузу 5 тонн", "Ташкент-Самарканд - Исузу 5тонн", 2420000, 165000),
    (9, "Самарканд", "10 тонн", "Ташкен-Самарканд - 10 тонн", 2970000, 220000),
    (10, "Фергана", "Газель", "Ташкент-Фергана - Газель", 2300000, 165000),
    (11, "Фергана", "Исузу 5 тонн", "Ташкент-Фергана - Исузу 5тонн", 2900000, 165000),
    (12, "Фергана", "10 тонн", "Ташкент-Фергана - 10 тонн", 3900000, 220000),
    (13, "Наманган", "Исузу 5 тонн", "Ташкент-Наманган - Исузу 5тонн", 2970000, 165000),
    (14, "Наманган", "10 тонн", "Ташкент-Наманган - 10 тонн", 3960000, 220000),
    (15, "Андижан", "Газель", "Ташкент-Андижан - Газель", 2600000, 165000),
    (16, "Андижан", "Исузу 5 тонн", "Ташкент-Андижан - Исузу 5тонн", 3300000, 220000),
    (17, "Навои", "2 тонны", "Ташкент-Навои 2 тонны", 2700000, 220000),
    (18, "Навои", "Исузу 5 тонн", "Ташкент-Навои 5 тонн", 4400000, 220000),
    (19, "Навои", "10 тонн", "Ташкент-Навои 10 тонн", 5900000, 220000),
    (20, "Бухара", "2 тонны", "Ташкент-Бухара 2 тонны", 3950000, 220000),
    (21, "Бухара", "Исузу 5 тонн", "Ташкент-Бухара 5 тонн", 4650000, 220000),
    (22, "Бухара", "10 тонн", "Ташкент-Бухара 10 тонн", 6500000, 220000),
    (23, None, None, "Дополнительно грузчик-экспедитор в машину (9:00-18:00)", 440000, None),
]

_TARIFF_COLS = ["№ п/п", "Регион", "Авто", "Маршрут", "Базовая ставка за рейс", "Стоимость дополнительной точки"]


def default_tariffs_df() -> pd.DataFrame:
    return pd.DataFrame(DEFAULT_TARIFFS, columns=_TARIFF_COLS)


def _serial_to_datetime(v):
    """Преобразует порядковый номер даты (Google Sheets/Excel) в datetime."""
    try:
        return _EXCEL_EPOCH + timedelta(days=float(v))
    except (TypeError, ValueError):
        return None


class _Cell:
    __slots__ = ("value",)

    def __init__(self, value):
        self.value = value


class ListWorksheet:
    """
    Адаптер: превращает список списков (сырые значения, например от gspread)
    в объект с интерфейсом openpyxl Worksheet (.cell(row, column).value,
    .max_row), чтобы parse_shipments/parse_orders_status работали одинаково
    и с .xlsx, и с данными из Google Sheets API.
    """

    def __init__(self, rows, date_columns=None):
        self.rows = rows
        self.date_columns = date_columns or set()
        self.max_row = len(rows)

    def cell(self, row, column):
        r, c = row - 1, column - 1
        if r < 0 or r >= len(self.rows):
            return _Cell(None)
        row_data = self.rows[r]
        if c < 0 or c >= len(row_data):
            return _Cell(None)
        v = row_data[c]
        if v == "":
            v = None
        if v is not None and column in self.date_columns and isinstance(v, (int, float)) and not isinstance(v, bool):
            dt = _serial_to_datetime(v)
            if dt is not None:
                v = dt
        return _Cell(v)


def _is_error(v):
    if v is None:
        return True
    if isinstance(v, str) and v.strip().startswith("#"):
        return True
    return False


def _num(v):
    if _is_error(v):
        return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _clean_str(v):
    if _is_error(v):
        return None
    if isinstance(v, str):
        v = v.strip()
        return v if v else None
    return v


def parse_shipments(ws) -> pd.DataFrame:
    """Лист 'Статус по отгрузкам' -> tidy DataFrame, 1 строка = 1 заказ в рейсе."""
    rows = []
    current_day = None
    current_reys = None
    for r in range(2, ws.max_row + 1):
        b = ws.cell(row=r, column=2).value  # Дата отгрузки
        c = ws.cell(row=r, column=3).value  # Заказ
        if b is None and c is None:
            continue
        if isinstance(b, str) and b.strip().lower().startswith("рейс"):
            current_reys = b.strip()
            continue
        if isinstance(b, datetime) and c is None:
            current_day = b.date()
            current_reys = None
            continue
        if c is not None:
            day_val = b.date() if isinstance(b, datetime) else current_day
            rows.append(
                {
                    "Дата": day_val,
                    "Рейс": current_reys,
                    "Заказ": str(c).strip(),
                    "Кол_во_шт": _num(ws.cell(row=r, column=4).value),
                    "ID_магазина": _clean_str(ws.cell(row=r, column=5).value),
                    "Магазин": _clean_str(ws.cell(row=r, column=6).value),
                    "Адрес": _clean_str(ws.cell(row=r, column=7).value),
                    "Координаты_1": _clean_str(ws.cell(row=r, column=8).value),
                    "Координаты_2": _clean_str(ws.cell(row=r, column=9).value),
                    "Регион": _clean_str(ws.cell(row=r, column=10).value),
                    "Транспорт": _clean_str(ws.cell(row=r, column=11).value),
                    "Точка_1": _num(ws.cell(row=r, column=12).value),
                    "Точка_2_и_далее": _num(ws.cell(row=r, column=13).value),
                    "Грузчик_экспедитор": _num(ws.cell(row=r, column=14).value),
                    "Итого_сумма": _num(ws.cell(row=r, column=15).value),
                    "Подрядчик": _clean_str(ws.cell(row=r, column=16).value),
                }
            )
    df = pd.DataFrame(rows)
    if not df.empty:
        df["Дата"] = pd.to_datetime(df["Дата"])
    return df


def parse_orders_status(ws) -> pd.DataFrame:
    """Лист 'Статус по заказам' -> tidy DataFrame с плановой датой отгрузки."""
    rows = []
    current_section_date = None
    for r in range(2, ws.max_row + 1):
        a = ws.cell(row=r, column=1).value
        if a is None:
            continue
        a_str = str(a).strip()
        if not a_str.startswith("1M-"):
            m = _DATE_RE.search(a_str)
            if m:
                d, mo, y = m.groups()
                try:
                    current_section_date = date(int(y), int(mo), int(d))
                except ValueError:
                    pass
            continue
        plan_cell = ws.cell(row=r, column=6).value  # Дата отгрузки план
        plan_date = plan_cell.date() if isinstance(plan_cell, datetime) else current_section_date
        rows.append(
            {
                "Заказ": a_str,
                "ID_магазина": _clean_str(ws.cell(row=r, column=2).value),
                "Магазин": _clean_str(ws.cell(row=r, column=3).value),
                "Адрес_магазина": _clean_str(ws.cell(row=r, column=4).value),
                "Кол_во_шт_заказ": _num(ws.cell(row=r, column=5).value),
                "Дата_план": plan_date,
                "Город": _clean_str(ws.cell(row=r, column=7).value),
                "Статус_сборки": _clean_str(ws.cell(row=r, column=8).value),
                "Статус_WMS": _clean_str(ws.cell(row=r, column=9).value),
                "Статус_отгрузки_на_хаб": _clean_str(ws.cell(row=r, column=10).value),
            }
        )
    df = pd.DataFrame(rows)
    if not df.empty:
        df["Дата_план"] = pd.to_datetime(df["Дата_план"])
    return df


def add_route_economics(ship_df: pd.DataFrame, region_tariff_map: dict = None) -> pd.DataFrame:
    """
    Пересчитывает стоимость доставки по логике заказчика и добавляет поля,
    нужные для биллинг-отчёта.

    В исходнике "Итого сумма" — это тариф за точку маршрута (магазин), который
    проставлен только один раз на уникальный магазин в рейсе (повторные заказы
    в тот же магазин в тот же рейс получают 0, доставка уже учтена). Из-за
    этого при группировке "по магазину" суммы получаются неровными — то 0, то
    полный тариф.

    Правильный расчёт:
        1. Сумма_маршрута = сумма "Итого сумма" по всем строкам рейса
           (это и есть полная стоимость доставки рейса — сумма тарифов по
           каждой уникальной точке, т.к. повторы уже дают 0).
        2. Кол_во_шт_маршрута = сумма "Кол-во шт" по всем строкам рейса
           (включая повторные заказы в тот же магазин — они тоже везли товар).
        3. Тариф_за_шт = Сумма_маршрута / Кол_во_шт_маршрута.
        4. Для каждой строки (заказа): Сумма_распределённая = Кол-во_шт
           этой строки * Тариф_за_шт — то есть стоимость доставки размазана
           по всем штукам маршрута пропорционально, а не только на первую
           точку.

    Дополнительно (для биллинг-реестра, формат как в примере заказчика):
        - Маршрут_ID — идентификатор рейса вида "1M-00xxxxx" (номер первого
          заказа в рейсе), уникальный в пределах всей выгрузки (в отличие от
          "Рейс 1"/"Рейс 2", которые каждый день начинаются заново).
        - Регион_тариф — регион, нормализованный к тарифной зоне (например,
          Янгиюль/Ангрен/Чирчик -> Ташкент, Коканд -> Фергана), по словарю
          region_tariff_map (по умолчанию DEFAULT_REGION_TARIFF_MAP).

    Группировка — по (Дата, Рейс), т.к. номера рейсов ("Рейс 1", "Рейс 2"...)
    повторяются день в день и относятся к разным маршрутам.

    Если внутри одного (Дата, Рейс) встречаются заказы из РАЗНЫХ городов
    (регионов) — это ошибка в исходнике (один развоз физически не может
    быть одновременно, например, в Коканде и в Намангане). В этом случае
    рейс целиком разбивается на отдельные маршруты — по одному заказу на
    маршрут, каждый со своим Маршрут_ID (= номер этого же заказа) и своей
    экономикой (без размазывания тарифа между заказами из разных городов).
    Если же город в рейсе один — группировка остаётся прежней: весь рейс
    считается одним маршрутом, тариф размазывается по всем заказам рейса
    пропорционально кол-ву штук, как и раньше.
    """
    if ship_df.empty:
        return ship_df

    region_map = region_tariff_map if region_tariff_map is not None else DEFAULT_REGION_TARIFF_MAP

    df = ship_df.copy()
    base_cols = ["Дата", "Рейс"]

    # Рейс, где перепутаны города, -> у каждого заказа в нём свой отдельный
    # маршрут (ключ = сам заказ); иначе -> общий ключ (Дата, Рейс), как и
    # для обычных рейсов.
    n_cities = df.groupby(base_cols, dropna=False)["Регион"].transform("nunique")
    multi_city = n_cities > 1
    df["_route_key"] = df["Заказ"].where(multi_city, df["Рейс"])

    grp_cols = ["Дата", "_route_key"]
    route_totals = (
        df.groupby(grp_cols, dropna=False)
        .agg(Сумма_маршрута=("Итого_сумма", "sum"), Кол_во_шт_маршрута=("Кол_во_шт", "sum"))
        .reset_index()
    )
    route_totals["Тариф_за_шт"] = route_totals.apply(
        lambda r: (r["Сумма_маршрута"] / r["Кол_во_шт_маршрута"]) if r["Кол_во_шт_маршрута"] else 0.0,
        axis=1,
    )
    df = df.merge(route_totals, on=grp_cols, how="left")
    df["Сумма_распределенная"] = df["Кол_во_шт"] * df["Тариф_за_шт"]

    # Маршрут_ID = номер первого заказа, встреченного в группе (обычный
    # рейс), либо номер самого заказа (если рейс был разбит из-за разных
    # городов — см. выше).
    anchor = df.groupby(grp_cols, dropna=False, sort=False)["Заказ"].transform("first")
    df["Маршрут_ID"] = anchor
    df = df.drop(columns=["_route_key"])

    df["Регион_тариф"] = df["Регион"].map(lambda r: region_map.get(r, r) if r is not None else r)

    return df


def parse_tariffs(ws) -> pd.DataFrame:
    """Лист 'тарифы' (справочник ставок) -> DataFrame с теми же заголовками."""
    rows = []
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, 7)]
        if all(v is None for v in vals):
            continue
        rows.append(vals)
    if not rows:
        return pd.DataFrame(columns=_TARIFF_COLS)
    return pd.DataFrame(rows, columns=_TARIFF_COLS)


def parse_workbook_bytes(data: bytes):
    """Возвращает (ship_df, orders_df) из байтов .xlsx файла."""
    wb = openpyxl.load_workbook(BytesIO(data), data_only=True)
    ship_df = parse_shipments(wb[SHIPMENTS_SHEET])
    orders_df = parse_orders_status(wb[ORDERS_SHEET])
    return ship_df, orders_df


_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]


def load_from_gsheet_service_account(sheet_id: str, credentials_info: dict):
    """
    Возвращает (ship_df, orders_df, tariffs_df), читая таблицу через Google
    Sheets API с помощью сервисного аккаунта. Таблицу не нужно открывать по
    ссылке — достаточно выдать доступ на чтение самому сервисному аккаунту
    (credentials_info['client_email']).

    Лист "тарифы" — необязательный: если такого листа в таблице нет,
    возвращается пустой DataFrame (в приложении в этом случае используется
    встроенный справочник по умолчанию, см. default_tariffs_df()).
    """
    import gspread
    from google.oauth2.service_account import Credentials

    creds = Credentials.from_service_account_info(dict(credentials_info), scopes=_SCOPES)
    client = gspread.authorize(creds)
    sh = client.open_by_key(sheet_id)

    ship_ws = sh.worksheet(SHIPMENTS_SHEET)
    orders_ws = sh.worksheet(ORDERS_SHEET)

    ship_rows = ship_ws.get_values(value_render_option="UNFORMATTED_VALUE")
    orders_rows = orders_ws.get_values(value_render_option="UNFORMATTED_VALUE")

    ship_adapter = ListWorksheet(ship_rows, date_columns=_SHIPMENTS_DATE_COLS)
    orders_adapter = ListWorksheet(orders_rows, date_columns=_ORDERS_DATE_COLS)

    ship_df = parse_shipments(ship_adapter)
    orders_df = parse_orders_status(orders_adapter)

    tariffs_df = pd.DataFrame(columns=_TARIFF_COLS)
    try:
        tariffs_ws = sh.worksheet(TARIFFS_SHEET)
        tariffs_rows = tariffs_ws.get_values(value_render_option="UNFORMATTED_VALUE")
        tariffs_adapter = ListWorksheet(tariffs_rows)
        tariffs_df = parse_tariffs(tariffs_adapter)
    except Exception:
        pass  # листа "тарифы" нет в таблице — это нормально, используем дефолт

    return ship_df, orders_df, tariffs_df


def build_sla(ship_df: pd.DataFrame, orders_df: pd.DataFrame) -> pd.DataFrame:
    """Джойн план (Статус по заказам) vs факт (Статус по отгрузкам) по номеру заказа 1M-."""
    if ship_df.empty or orders_df.empty:
        return pd.DataFrame()

    fact = (
        ship_df[ship_df["Заказ"].str.startswith("1M-")]
        .groupby("Заказ", as_index=False)
        .agg(Дата_факт=("Дата", "min"))
    )
    plan = orders_df[orders_df["Заказ"].str.startswith("1M-")][
        ["Заказ", "Магазин", "Город", "Дата_план", "Статус_отгрузки_на_хаб"]
    ]
    sla = plan.merge(fact, on="Заказ", how="left")
    sla["Дельта_дней"] = (sla["Дата_факт"] - sla["Дата_план"]).dt.days
    sla["SLA_статус"] = sla.apply(_sla_flag, axis=1)
    return sla


def _sla_flag(row):
    if pd.isna(row["Дата_план"]):
        return "Нет плана"
    if pd.isna(row["Дата_факт"]):
        return "Не отгружен"
    if row["Дельта_дней"] <= 0:
        return "В срок"
    return "Просрочка"


def attach_city(ship_df: pd.DataFrame, orders_df: pd.DataFrame) -> pd.DataFrame:
    """
    Добавляет колонку "Город" в ship_df: берётся из листа "Статус по заказам"
    (по номеру заказа), с фолбэком на колонку "Регион" из "Статус по
    отгрузкам", если заказ там не найден. Используется во всех местах, где
    нужен город — вкладка "Сводная", "Детали по магазинам", экспорт.
    """
    df = ship_df.copy()
    if orders_df is not None and not orders_df.empty and "Город" in orders_df.columns:
        df = df.merge(orders_df[["Заказ", "Город"]], on="Заказ", how="left")
    else:
        df["Город"] = None
    df["Город"] = df["Город"].fillna(df["Регион"])
    return df


# Формат дат во всех Excel-экспортах — ISO (2026-07-01), по просьбе заказчика.
_ISO_DATE_FMT = "yyyy-mm-dd"


def _prepare_registry_df(ship_df: pd.DataFrame) -> pd.DataFrame:
    """
    Готовит построчный реестр для экспорта: без "Подрядчик", с "Городом"
    (см. attach_city — должен быть уже вызван до этого), с датой как date
    (не datetime, чтобы в Excel не показывало "00:00:00").

    "Итого сумма" в реестре — это ЧЕСТНО РАСПРЕДЕЛЁННАЯ сумма
    (Сумма_распределенная, см. add_route_economics). Исходный тариф из
    таблицы (до распределения) остаётся отдельной колонкой
    "Тариф (исходный)" — для сверки/аудита.
    """
    df = ship_df.sort_values(["Дата", "Маршрут_ID", "Заказ"]).copy()
    out = pd.DataFrame(
        {
            "Дата": df["Дата"].dt.date,
            "Заказ": df["Заказ"],
            "Рейс": df["Маршрут_ID"],
            "Кол-во шт": df["Кол_во_шт"],
            "ID магазина": df["ID_магазина"],
            "Магазин": df["Магазин"],
            "Адрес": df["Адрес"],
            "Координаты 1": df["Координаты_1"],
            "Координаты 2": df["Координаты_2"],
            "Регион": df["Регион"],
            "Регион тариф": df["Регион_тариф"],
            "Город": df["Город"] if "Город" in df.columns else df["Регион"],
            "Авто": df["Транспорт"],
            "1 точка": df["Точка_1"],
            "2 точка и далее": df["Точка_2_и_далее"],
            "Грузчик экспедитор": df["Грузчик_экспедитор"],
            "Тариф (исходный)": df["Итого_сумма"],
            "Итого сумма": df["Сумма_распределенная"],
        }
    )
    return out.reset_index(drop=True)


def build_registry_export(ship_df: pd.DataFrame) -> bytes:
    """
    Excel-экспорт — ровно один лист "Реестр", построчно по каждому заказу
    (оформлен как Excel-таблица с фильтрами). Без "Подрядчик". Даты — в
    формате 2026-07-01. Остальные сводные вкладки (по решению заказчика) в
    файл не включаются — сводки считаются в самом приложении (вкладки
    "Отгрузки по дням" / "Сводная"), а не в выгрузке.

    ship_df должен быть уже отфильтрован (период) и содержать колонки,
    которые добавляют add_route_economics() и attach_city().
    """
    from io import BytesIO as _BytesIO

    import openpyxl as _openpyxl
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo

    wb = _openpyxl.Workbook()
    ws = wb.active
    ws.title = "Реестр"

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4C78A8")

    registry = _prepare_registry_df(ship_df)
    registry_cols = list(registry.columns)
    date_col_idx = registry_cols.index("Дата") + 1  # 1-based

    ws.append(registry_cols)
    for c in range(1, len(registry_cols) + 1):
        ws.cell(row=1, column=c).font = header_font
        ws.cell(row=1, column=c).fill = header_fill

    for row in registry.itertuples(index=False):
        ws.append(list(row))

    last_row = 1 + len(registry)
    if last_row > 1:
        table_ref = f"A1:{get_column_letter(len(registry_cols))}{last_row}"
        tbl = Table(displayName="ReestrTable", ref=table_ref)
        tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(tbl)
        for r in range(2, last_row + 1):
            ws.cell(row=r, column=date_col_idx).number_format = _ISO_DATE_FMT

    for c in range(1, len(registry_cols) + 1):
        letter = get_column_letter(c)
        max_len = 10
        for row in ws.iter_rows(min_col=c, max_col=c):
            v = row[0].value
            if v is not None:
                max_len = max(max_len, len(str(v)))
        ws.column_dimensions[letter].width = min(max_len + 2, 42)
    ws.freeze_panes = "A2"

    buf = _BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_sla_export(sla_df: pd.DataFrame) -> bytes:
    """Отдельный Excel-файл с SLA (план vs факт) — не входит в основной отчёт."""
    from io import BytesIO as _BytesIO

    import openpyxl as _openpyxl
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = _openpyxl.Workbook()
    ws = wb.active
    ws.title = "SLA план-факт"

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4C78A8")

    if sla_df is None or sla_df.empty:
        ws.append(["Нет данных"])
    else:
        cols = list(sla_df.columns)
        ws.append(cols)
        for c in range(1, len(cols) + 1):
            ws.cell(row=1, column=c).font = header_font
            ws.cell(row=1, column=c).fill = header_fill
        date_cols = [i for i, c in enumerate(cols) if c.startswith("Дата_")]
        for _, row in sla_df.iterrows():
            vals = []
            for i, c in enumerate(cols):
                v = row[c]
                if i in date_cols:
                    v = v.date() if pd.notna(v) and hasattr(v, "date") else None
                vals.append(v)
            ws.append(vals)
        for i in date_cols:
            for r in range(2, ws.max_row + 1):
                ws.cell(row=r, column=i + 1).number_format = _ISO_DATE_FMT
        for c in range(1, len(cols) + 1):
            letter = get_column_letter(c)
            max_len = max(
                (len(str(ws.cell(row=r, column=c).value or "")) for r in range(1, ws.max_row + 1)),
                default=10,
            )
            ws.column_dimensions[letter].width = min(max_len + 2, 40)
        ws.freeze_panes = "A2"

    buf = _BytesIO()
    wb.save(buf)
    return buf.getvalue()
