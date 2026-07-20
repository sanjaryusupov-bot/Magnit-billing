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
    (не datetime), в порядке колонок, читаемом для сводной таблицы.

    "Итого сумма" в реестре — это ЧЕСТНО РАСПРЕДЕЛЁННАЯ сумма
    (Сумма_распределенная, см. add_route_economics), т.к. именно её нужно
    суммировать в отчётах/сводной таблице. Исходный тариф из таблицы (до
    распределения) остаётся отдельной колонкой "Тариф (исходный)" — для
    сверки/аудита.
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


def _add_native_pivot(
    xlsx_bytes: bytes,
    *,
    source_table: str,
    cache_field_names: list,
    row_field_idx: list,
    data_fields: list,  # [(cache_field_idx, "Название"), ...]
    target_sheet_name: str,
    field_num_fmt: dict = None,
    pivot_name: str = "SvodnayaPivot",
) -> bytes:
    """
    Дописывает в уже сохранённый .xlsx (bytes) НАСТОЯЩУЮ Excel-сводную
    таблицу (PivotTable), построенную поверх Excel-таблицы source_table
    (объект openpyxl.worksheet.table.Table на одном из листов книги).

    openpyxl умеет читать сводные таблицы, но не даёт высокоуровневого API
    для их создания — поэтому pivotCacheDefinition / pivotCacheRecords /
    pivotTable дописываются как отдельные XML-части вручную, напрямую в
    .xlsx (это ZIP-архив).

    Кэш создаётся с refreshOnLoad="1" и пустыми записями (recordCount="0") —
    это стандартный приём: Excel сам пересчитывает содержимое сводной из
    исходной таблицы при открытии файла, поэтому нам не нужно самим
    формировать корректные значения кэша построчно. Проверено:
    результат открывается и корректно пересчитывается (сверено на реальных
    суммах) в LibreOffice, которая гораздо строже парсит OOXML, чем Excel.

    Это самый настоящий PivotTable — двойной клик по значению в Excel
    покажет исходные строки реестра, из которых оно посчитано (стандартное
    поведение Excel "Показать сведения", ничего специально настраивать не
    нужно — оно работает у любой корректной сводной таблицы).
    """
    import re as _re
    import zipfile as _zipfile
    from io import BytesIO as _BytesIO

    field_num_fmt = field_num_fmt or {}

    zin = _zipfile.ZipFile(_BytesIO(xlsx_bytes), "r")
    parts = {name: zin.read(name) for name in zin.namelist()}
    zin.close()

    wbxml = parts["xl/workbook.xml"].decode("utf-8")
    m = _re.search(
        r'<sheet[^>]*\bname="%s"[^>]*\br:id="([^"]+)"' % _re.escape(target_sheet_name), wbxml
    )
    if not m:
        raise ValueError(f"Лист '{target_sheet_name}' не найден в workbook.xml")
    sheet_rid = m.group(1)

    wbrels = parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
    m2 = _re.search(r'Id="%s"[^>]*Target="([^"]+)"' % _re.escape(sheet_rid), wbrels) or _re.search(
        r'Target="([^"]+)"[^>]*Id="%s"' % _re.escape(sheet_rid), wbrels
    )
    if not m2:
        raise ValueError("Не удалось найти relationship листа")
    sheet_target = m2.group(1).lstrip("/")
    if not sheet_target.startswith("xl/"):
        sheet_target = "xl/" + sheet_target
    sheet_file = sheet_target.rsplit("/", 1)[-1]
    sheet_rels_path = f"xl/worksheets/_rels/{sheet_file}.rels"

    n_fields = len(cache_field_names)

    cache_fields_xml = "".join(
        f'<cacheField name="{name}" numFmtId="{field_num_fmt.get(i, 0)}"><sharedItems/></cacheField>'
        for i, name in enumerate(cache_field_names)
    )
    cache_def = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'refreshOnLoad="1" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="0">'
        f'<cacheSource type="worksheet"><worksheetSource name="{source_table}"/></cacheSource>'
        f'<cacheFields count="{n_fields}">{cache_fields_xml}</cacheFields>'
        "</pivotCacheDefinition>"
    )
    cache_records = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" count="0"/>'
    )
    cache_def_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" '
        'Target="pivotCacheRecords1.xml"/></Relationships>'
    )

    data_idx = {idx for idx, _ in data_fields}
    row_set = set(row_field_idx)
    pivot_fields_xml = "".join(
        '<pivotField dataField="1" showAll="0"/>' if i in data_idx
        else '<pivotField axis="axisRow" showAll="0"/>' if i in row_set
        else '<pivotField showAll="0"/>'
        for i in range(n_fields)
    )

    row_fields_xml = "".join(f'<field x="{i}"/>' for i in row_field_idx)
    n_row_fields = len(row_field_idx)
    n_data_fields = len(data_fields)
    data_fields_xml = "".join(
        f'<dataField name="{label}" fld="{idx}" baseField="0" baseItem="0"/>' for idx, label in data_fields
    )

    if n_data_fields >= 2:
        col_fields_xml = '<colFields count="1"><field x="-2"/></colFields>'
        col_items = ["<i><x/></i>"] + [f'<i i="{k}"><x v="{k}"/></i>' for k in range(1, n_data_fields)]
        col_items_xml = f'<colItems count="{n_data_fields}">' + "".join(col_items) + "</colItems>"
    else:
        col_fields_xml = ""
        col_items_xml = '<colItems count="1"><i/></colItems>'

    pivot_table = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'name="{pivot_name}" cacheId="1" applyNumberFormats="0" applyBorderFormats="0" '
        'applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" '
        'dataCaption="Значения" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" '
        'itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">'
        f'<location ref="A1:B2" firstHeaderRow="1" firstDataRow="2" firstDataCol="{n_row_fields}"/>'
        f'<pivotFields count="{n_fields}">{pivot_fields_xml}</pivotFields>'
        f'<rowFields count="{n_row_fields}">{row_fields_xml}</rowFields>'
        '<rowItems count="1"><i><x/></i></rowItems>'
        f"{col_fields_xml}{col_items_xml}"
        f'<dataFields count="{n_data_fields}">{data_fields_xml}</dataFields>'
        '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" '
        'showRowStripes="0" showColStripes="0" showLastColumn="1"/>'
        "</pivotTableDefinition>"
    )
    pivot_table_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" '
        'Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>'
    )

    # pivotCaches должен идти ПОСЛЕ <calcPr/> — того требует порядок элементов
    # в схеме CT_Workbook; иначе Excel/LibreOffice не могут открыть файл.
    wbxml, n = _re.subn(
        r"(<calcPr[^>]*/>)",
        r'\1<pivotCaches><pivotCache cacheId="1" '
        r'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        r'r:id="rIdPivotCacheDef1"/></pivotCaches>',
        wbxml,
    )
    if n != 1:
        raise ValueError("Не найден <calcPr/> в workbook.xml — не удалось встроить сводную таблицу")
    parts["xl/workbook.xml"] = wbxml.encode("utf-8")

    wbrels = wbrels.replace(
        "</Relationships>",
        '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" '
        'Target="pivotCache/pivotCacheDefinition1.xml" Id="rIdPivotCacheDef1"/></Relationships>',
    )
    parts["xl/_rels/workbook.xml.rels"] = wbrels.encode("utf-8")

    if sheet_rels_path in parts:
        r = parts[sheet_rels_path].decode("utf-8").replace(
            "</Relationships>",
            '<Relationship Id="rIdPivotTable1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" '
            'Target="../pivotTables/pivotTable1.xml"/></Relationships>',
        )
    else:
        r = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rIdPivotTable1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" '
            'Target="../pivotTables/pivotTable1.xml"/></Relationships>'
        )
    parts[sheet_rels_path] = r.encode("utf-8")

    parts["xl/pivotCache/pivotCacheDefinition1.xml"] = cache_def.encode("utf-8")
    parts["xl/pivotCache/pivotCacheRecords1.xml"] = cache_records.encode("utf-8")
    parts["xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"] = cache_def_rels.encode("utf-8")
    parts["xl/pivotTables/pivotTable1.xml"] = pivot_table.encode("utf-8")
    parts["xl/pivotTables/_rels/pivotTable1.xml.rels"] = pivot_table_rels.encode("utf-8")

    ct = parts["[Content_Types].xml"].decode("utf-8").replace(
        "</Types>",
        '<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>'
        '<Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>'
        '<Override PartName="/xl/pivotTables/pivotTable1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>'
        "</Types>",
    )
    parts["[Content_Types].xml"] = ct.encode("utf-8")

    out = _BytesIO()
    with _zipfile.ZipFile(out, "w", _zipfile.ZIP_DEFLATED) as zout:
        for name, data in parts.items():
            zout.writestr(name, data)
    return out.getvalue()


def build_pivot_export(ship_df: pd.DataFrame, tariffs_df: pd.DataFrame) -> bytes:
    """
    Основной Excel-экспорт с НАСТОЯЩЕЙ сводной таблицей (PivotTable):

      - "Реестр" — построчно по каждому заказу (Excel-таблица, источник для
        сводной), без колонки "Подрядчик".
      - "Сводная" — живая PivotTable: Регион тариф -> Транспорт -> Рейс ->
        Дата, с суммами "Итого сумма" (честно распределённая) и "Кол-во шт".
        Двойной клик по любой цифре в Excel открывает лист с заказами,
        из которых она посчитана (родное поведение Excel).
      - "Сводная по дням" — Дата отгрузки, Итого сумма, Кол-во штук, Сумма за шт.
      - "тарифы" — справочник ставок.

    ship_df должен быть уже отфильтрован (период) и содержать колонки,
    которые добавляют add_route_economics() и attach_city().
    """
    from io import BytesIO as _BytesIO

    import openpyxl as _openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo

    wb = _openpyxl.Workbook()

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4C78A8")
    bold = Font(bold=True)

    def _style_header(ws, row, ncols):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=row, column=c)
            cell.font = header_font
            cell.fill = header_fill

    def _autosize(ws, ncols, min_width=10, max_width=42):
        for c in range(1, ncols + 1):
            letter = get_column_letter(c)
            max_len = min_width
            for row in ws.iter_rows(min_col=c, max_col=c):
                v = row[0].value
                if v is not None:
                    max_len = max(max_len, len(str(v)))
            ws.column_dimensions[letter].width = min(max_len + 2, max_width)

    # ---------------- Лист "Реестр" (источник для сводной) ----------------
    ws_reg = wb.active
    ws_reg.title = "Реестр"
    registry = _prepare_registry_df(ship_df)
    registry_cols = list(registry.columns)
    date_col_idx = registry_cols.index("Дата") + 1  # 1-based

    ws_reg.append(registry_cols)
    for row in registry.itertuples(index=False):
        ws_reg.append(list(row))

    last_row = 1 + len(registry)
    if last_row > 1:
        table_ref = f"A1:{get_column_letter(len(registry_cols))}{last_row}"
        tbl = Table(displayName="ReestrTable", ref=table_ref)
        tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws_reg.add_table(tbl)
        for r in range(2, last_row + 1):
            ws_reg.cell(row=r, column=date_col_idx).number_format = _ISO_DATE_FMT
    _autosize(ws_reg, len(registry_cols), max_width=38)

    # numFmtId, который openpyxl присвоил формату ISO-даты (нужен для того,
    # чтобы даты в самой сводной тоже отображались как 2026-07-01, а не
    # как порядковый номер или полный datetime).
    date_fmt_id = 0
    if last_row > 1:
        buf_probe = _BytesIO()
        wb.save(buf_probe)
        m = re.search(r'numFmtId="(\d+)" formatCode="yyyy-mm-dd"', buf_probe.getvalue().decode("latin1"))
        if m:
            date_fmt_id = int(m.group(1))

    # ---------------- Лист "Сводная" (плейсхолдер под живую PivotTable) ----------------
    ws_pivot = wb.create_sheet("Сводная")

    # ---------------- Лист "Сводная по дням" ----------------
    ws_daily = wb.create_sheet("Сводная по дням")
    ws_daily.append(["Дата отгрузки", "Итого сумма", "Кол-во штук", "Сумма за шт"])
    _style_header(ws_daily, 1, 4)
    if not ship_df.empty:
        daily = (
            ship_df.groupby("Дата", as_index=False)
            .agg(Сумма=("Сумма_распределенная", "sum"), Кол_во=("Кол_во_шт", "sum"))
            .sort_values("Дата")
        )
        for _, row in daily.iterrows():
            rate = (row["Сумма"] / row["Кол_во"]) if row["Кол_во"] else 0.0
            ws_daily.append([row["Дата"].date(), row["Сумма"], row["Кол_во"], round(rate, 2)])
        total_row = ws_daily.max_row + 1
        total_sum, total_qty = daily["Сумма"].sum(), daily["Кол_во"].sum()
        total_rate = (total_sum / total_qty) if total_qty else 0.0
        ws_daily.append(["Итого", total_sum, total_qty, round(total_rate, 2)])
        for c in range(1, 5):
            ws_daily.cell(row=total_row, column=c).font = bold
    for col, width in zip("ABCD", [16, 16, 14, 14]):
        ws_daily.column_dimensions[col].width = width
    for row in ws_daily.iter_rows(min_row=2, max_row=ws_daily.max_row, min_col=1, max_col=1):
        row[0].number_format = _ISO_DATE_FMT

    # ---------------- Лист "тарифы" ----------------
    ws_tar = wb.create_sheet("тарифы")
    t_df = tariffs_df if (tariffs_df is not None and not tariffs_df.empty) else default_tariffs_df()
    ws_tar.append(list(t_df.columns))
    _style_header(ws_tar, 1, len(t_df.columns))
    for _, row in t_df.iterrows():
        ws_tar.append(list(row.values))
    _autosize(ws_tar, len(t_df.columns), max_width=60)

    buf = _BytesIO()
    wb.save(buf)
    base_bytes = buf.getvalue()

    if last_row <= 1:
        # Нет данных за период — отдаём книгу без сводной (её не из чего строить).
        return base_bytes

    # Регион тариф, Транспорт, Рейс, Дата — индексы в registry_cols (0-based).
    row_field_idx = [
        registry_cols.index("Регион тариф"),
        registry_cols.index("Авто"),
        registry_cols.index("Рейс"),
        registry_cols.index("Дата"),
    ]
    data_fields = [
        (registry_cols.index("Итого сумма"), "Сумма по полю Итого сумма"),
        (registry_cols.index("Кол-во шт"), "Сумма по полю Кол-во шт"),
    ]
    field_num_fmt = {registry_cols.index("Дата"): date_fmt_id}

    return _add_native_pivot(
        base_bytes,
        source_table="ReestrTable",
        cache_field_names=registry_cols,
        row_field_idx=row_field_idx,
        data_fields=data_fields,
        target_sheet_name="Сводная",
        field_num_fmt=field_num_fmt,
    )


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
                if i in date_cols and pd.notna(v):
                    v = v.date() if hasattr(v, "date") else v
                elif i in date_cols:
                    v = None
                vals.append(v)
            ws.append(vals)
        for i in date_cols:
            letter = get_column_letter(i + 1)
            for r in range(2, ws.max_row + 1):
                ws.cell(row=r, column=i + 1).number_format = _ISO_DATE_FMT
        for c in range(1, len(cols) + 1):
            letter = get_column_letter(c)
            max_len = max((len(str(ws.cell(row=r, column=c).value or "")) for r in range(1, ws.max_row + 1)), default=10)
            ws.column_dimensions[letter].width = min(max_len + 2, 40)

    buf = _BytesIO()
    wb.save(buf)
    return buf.getvalue()
