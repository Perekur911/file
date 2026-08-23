from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class ColumnSchema:
    """
    Описание одного столбца таблицы.

    db_name:
        Имя столбца в SQLite.
    excel_name:
        Имя столбца в умной таблице Excel.
    sqlite_type:
        Тип хранения в SQLite: INTEGER, REAL, TEXT.
    value_type:
        Логический тип значения для конвертации:
        text, integer, real, boolean, date, time, datetime.
    primary_key:
        True для числового id таблицы.
    required:
        True, если поле обязательно при сохранении Excel -> БД.
    readonly:
        True для полей, которые не нужно писать в INSERT/UPDATE.
    """

    db_name: str
    excel_name: str | None = None
    sqlite_type: str = "TEXT"
    value_type: str = "text"
    primary_key: bool = False
    required: bool = False
    readonly: bool = False

    @property
    def xlsx_name(self) -> str:
        """Имя столбца в Excel."""
        return self.excel_name or self.db_name


@dataclass(frozen=True)
class TableSchema:
    """
    Описание соответствия одной умной таблицы Excel и одной таблицы SQLite.

    logical_name:
        Внутреннее имя таблицы, например AP_RESULTS.
    db_table:
        Имя таблицы в SQLite.
    sheet_name:
        Имя листа Excel, где лежит умная таблица.
    excel_table_name:
        Имя умной таблицы Excel. Если None, берётся первая умная таблица на листе.
    columns:
        Описание всех столбцов. Порядок в Excel может быть любым.
    parent:
        logical_name родительской таблицы для дочерних таблиц.
    fk_column:
        Колонка дочерней таблицы, где хранится id родительской строки.
    natural_key:
        Резерв на будущее для поиска строк без id по естественному ключу.
    """

    logical_name: str
    db_table: str
    sheet_name: str
    columns: list[ColumnSchema]
    excel_table_name: str | None = None
    parent: str | None = None
    fk_column: str | None = None
    natural_key: tuple[str, ...] = field(default_factory=tuple)
    # Колонки, из которых берём дату/время эксперимента для TaskStatus.
    # Можно указать одну колонку: ("date",)
    # Или две: ("dateStart", "timeStart")
    # Или приоритетный список: ("dateEnd", "timeEnd", "dateStart", "timeStart")
    status_datetime_fields: tuple[str, ...] = field(default_factory=tuple)
    delete_if_empty_columns: tuple[str, ...] = field(default_factory=tuple)

    @property
    def pk(self) -> ColumnSchema:
        """Единственный primary key таблицы."""
        keys = [c for c in self.columns if c.primary_key]
        if len(keys) != 1:
            raise ValueError(f"Для таблицы {self.logical_name} должен быть ровно один primary_key")
        return keys[0]

    @property
    def db_columns(self) -> list[str]:
        """Все столбцы БД в порядке описания схемы."""
        return [c.db_name for c in self.columns]

    @property
    def excel_columns(self) -> list[str]:
        """Все ожидаемые столбцы Excel в порядке описания схемы."""
        return [c.xlsx_name for c in self.columns]

    @property
    def excel_to_db(self) -> dict[str, str]:
        """Маппинг Excel-заголовок -> SQLite-столбец."""
        return {c.xlsx_name: c.db_name for c in self.columns}

    @property
    def db_to_excel(self) -> dict[str, str]:
        """Маппинг SQLite-столбец -> Excel-заголовок."""
        return {c.db_name: c.xlsx_name for c in self.columns}

    def insert_columns(self) -> list[str]:
        """Столбцы для INSERT. id не вставляется, если строка новая."""
        return [c.db_name for c in self.columns if not c.primary_key and not c.readonly]

    def update_columns(self) -> list[str]:
        """Столбцы для UPDATE. id и readonly-поля не обновляются."""
        return [c.db_name for c in self.columns if not c.primary_key and not c.readonly]


@dataclass(frozen=True)
class StudySchema:
    """
    Описание одного типа исследования.

    progress_step_field:
        Имя колонки дочерней таблицы, по которой рассчитывается прогресс.

    progress_total_steps:
        Число обычных этапов, отображаемое в статусе:
        "В работе N/progress_total_steps".

    completion_step:
        Специальный финальный этап. Его наличие означает "Завершено".
    """

    code: str
    result_table: TableSchema
    source_table: TableSchema | None = None

    progress_step_field: str | None = None
    progress_total_steps: int | None = None
    completion_step: int | None = None

    @property
    def has_progress_status(self) -> bool:
        return (
            self.source_table is not None
            and self.progress_step_field is not None
            and self.progress_total_steps is not None
            and self.completion_step is not None
        )


def _col(
    db_name: str,
    excel_name: str | None = None,
    typ: str = "TEXT",
    *,
    value_type: str = "text",
    pk: bool = False,
    required: bool = False,
    readonly: bool = False,
) -> ColumnSchema:
    """Короткий helper для компактного описания таблиц."""
    return ColumnSchema(
        db_name=db_name,
        excel_name=excel_name,
        sqlite_type=typ,
        value_type=value_type,
        primary_key=pk,
        required=required,
        readonly=readonly,
    )


# =============================================================================
# Описание таблиц.
#
# Имена столбцов в Excel и SQLite одинаковые, поэтому excel_name не указывается.
# Порядок столбцов в Excel может отличаться от порядка ниже.
# =============================================================================

OP_RESULTS = TableSchema(
    logical_name="OP_RESULTS",
    db_table="OP_results",
    sheet_name="OP_results",
    excel_table_name="OP_results",
    status_datetime_fields=("date",),
    columns=[
        _col("resultIdOP", "resultIdOP", typ="INTEGER", value_type="integer", pk=True),
        _col("TaskId", "TaskId", typ="INTEGER", value_type="integer"),
        _col("sampleCode", "sampleCode", typ="TEXT", value_type="text", required=True),
        _col("operator", "operator", typ="TEXT", value_type="text"),
        _col("date", "date", typ="TEXT", value_type="date"),
        _col("Popen", "Popen", typ="REAL", value_type="real"),
        _col("Punit", "Punit", typ="TEXT", value_type="text"),
        _col("Pabsolute", "Pabsolute", typ="INTEGER", value_type="boolean"),
        _col("PopenMPa", "PopenMPa", typ="REAL", value_type="real"),
        _col("Topen", "Topen", typ="REAL", value_type="real"),
        _col("VH2O", "VH2O", typ="REAL", value_type="real"),
        _col("Pend", "Pend", typ="REAL", value_type="real"),
        _col("PendMPa", "PendMPa", typ="REAL", value_type="real"),
        _col("VLiqPhase", "VLiqPhase", typ="REAl", value_type="real"),
        _col("LiqPhaseQualifier", "LiqPhaseQualifier", typ="TEXT", value_type="text"),
        _col("natureLiq", "natureLiq", typ="TEXT", value_type="text"),
        _col("state", "state", typ="INTEGER", value_type="boolean"),
        _col("comment", "comment", typ="TEXT", value_type="text"),
        _col("deltaPopenPCT", "deltaPopenPCT", typ="REAL", value_type="real"),
        _col("PopenMPaT", "PopenMPaT", typ="REAL", value_type="real"),
        _col("dateTimeSync", "dateTimeSync", typ="TEXT", value_type="datetime"),
    ],
)

OPOH_RESULTS = TableSchema(
    logical_name="OPOH_RESULTS",
    db_table="OPOH_results",
    sheet_name="OPOH_results",
    excel_table_name="OPOH_results",
    status_datetime_fields=("openDateTime",),
    columns=[
        _col("resultIdOPOH", "resultIdOPOH", typ="INTEGER", value_type="integer", pk=True),

        _col("TaskId", "TaskId", typ="INTEGER", value_type="integer"),
        _col("sampleCode", "sampleCode", typ="TEXT", value_type="text", required=True),

        _col("openDateTime", "openDateTime", typ="TEXT", value_type="datetime"),
        _col("valveState", "valveState", typ="INTEGER", value_type="boolean"),

        _col("TColdOpen", "TColdOpen", typ="REAL", value_type="real"),
        _col("coldCompP", "coldCompP", typ="REAL", value_type="real"),
        _col("coldPunit", "coldPunit", typ="TEXT", value_type="text"),
        _col("coldPabs", "coldPabs", typ="INTEGER", value_type="boolean"),
        _col("coldExpP", "coldExpP", typ="REAL", value_type="real"),
        _col("coldOpenPMPaAbs", "coldOpenPMPaAbs", typ="REAL", value_type="real"),

        _col("THotOpen", "THotOpen", typ="REAL", value_type="real"),
        _col("hotCompP", "hotCompP", typ="REAL", value_type="real"),
        _col("hotPunit", "hotPunit", typ="TEXT", value_type="text"),
        _col("hotPabs", "hotPabs", typ="INTEGER", value_type="boolean"),
        _col("hotExpP", "hotExpP", typ="REAL", value_type="real"),
        _col("hotOpenPMPaAbs", "hotOpenPMPaAbs", typ="REAL", value_type="real"),

        _col("openOperator", "openOperator", typ="TEXT", value_type="text"),

        _col("restorationP", "restorationP", typ="REAL", value_type="real"),
        _col("restorationPunit", "restorationPunit", typ="TEXT", value_type="text"),
        _col("restorationPabs", "restorationPabs", typ="INTEGER", value_type="boolean"),
        _col("restorationPMPaAbs", "restorationPMPaAbs", typ="REAL", value_type="real"),
        _col("restorationT", "restorationT", typ="REAL", value_type="real"),
        _col("restorationStartDateTime", "restorationStartDateTime", typ="TEXT", value_type="datetime"),
        _col("restorationStartOperator", "restorationStartOperator", typ="TEXT", value_type="text"),
        _col("restorationEndDateTime", "restorationEndDateTime", typ="TEXT", value_type="datetime"),
        _col("restorationEndOperator", "restorationEndOperator", typ="TEXT", value_type="text"),

        _col("typeSampler", "typeSampler", typ="TEXT", value_type="text"),
        _col("numberSampler", "numberSampler", typ="TEXT", value_type="text"),

        _col("transferP", "transferP", typ="REAL", value_type="real"),
        _col("transferPunit", "transferPunit", typ="TEXT", value_type="text"),
        _col("transferPabs", "transferPabs", typ="INTEGER", value_type="boolean"),
        _col("transferPMPaAbs", "transferPMPaAbs", typ="REAL", value_type="real"),
        _col("transferT", "transferT", typ="REAL", value_type="real"),
        _col("transferVolume", "transferVolume", typ="REAL", value_type="real"),
        _col("transferDateTime", "transferDateTime", typ="TEXT", value_type="datetime"),
        _col("transferOperator", "transferOperator", typ="TEXT", value_type="text"),

        _col("phaseDrainDateTime", "phaseDrainDateTime", typ="TEXT", value_type="datetime"),

        _col("sncVolumeHC", "sncVolumeHC", typ="REAL", value_type="real"),
        _col("sncVolumeDM", "sncVolumeDM", typ="REAL", value_type="real"),
        _col("sncVolumeW", "sncVolumeW", typ="REAL", value_type="real"),

        _col("flcVolumeHC", "flcVolumeHC", typ="REAL", value_type="real"),
        _col("flcVolumeDM", "flcVolumeDM", typ="REAL", value_type="real"),
        _col("flcVolumeW", "flcVolumeW", typ="REAL", value_type="real"),

        _col("phaseDrainOperator", "phaseDrainOperator", typ="TEXT", value_type="text"),
        _col("comment", "comment", typ="TEXT", value_type="text"),
        _col("dateTimeSync", "dateTimeSync", typ="TEXT", value_type="datetime"),
    ],
)

AP_RESULTS = TableSchema(
    logical_name="AP_RESULTS",
    db_table="AP_results",
    sheet_name="AP_results",
    status_datetime_fields=("date",),
    excel_table_name="AP_results",
    columns=[
        _col("resultIdAP", "resultIdAP", typ="INTEGER", value_type="integer", pk=True),
        _col("TaskId", "TaskId", typ="INTEGER", value_type="integer"),
        _col("sampleCode", "sampleCode", typ="TEXT", value_type="text", required=True),
        _col("date", "date", typ="TEXT", value_type="date"),
        _col("operator", "operator", typ="TEXT", value_type="text"),
        _col("analyseNumber", "analyseNumber", typ="TEXT", value_type="text"),
        _col("apEquipment", "apEquipment", typ="TEXT", value_type="text"),
        _col("apEquipmentNumber", "apEquipmentNumber", typ="TEXT", value_type="text"),
        _col("Tfact", "Tfact", typ="REAL", value_type="real"),
        _col("Correct", "Correct", typ="INTEGER", value_type="boolean"),
        _col("ka", "ka", typ="REAL", value_type="real"),
        _col("kb", "kb", typ="REAL", value_type="real"),
        _col("kc", "kc", typ="REAL", value_type="real"),
        _col("dateTimeSync", "dateTimeSync", typ="TEXT", value_type="datetime"),
    ],
)

AP_SOURCE_DATA = TableSchema(
    logical_name="AP_SOURCE_DATA",
    db_table="AP_sourceData",
    sheet_name="AP_sourceData",
    parent="AP_RESULTS",
    fk_column="resultIdAP",
    delete_if_empty_columns=("pressure", "density", "temperature"),
    excel_table_name="AP_sourceData",
    columns=[
        _col("rowIdAP", "rowIdAP", typ="INTEGER", value_type="integer", pk=True),
        _col("resultIdAP", "resultIdAP", typ="INTEGER", value_type="integer", required=True),
        _col("pressure", "pressure", typ="REAL", value_type="real"),
        _col("temperature", "temperature", typ="REAL", value_type="real"),
        _col("density", "density", typ="REAL", value_type="real"),
        _col("active", "active", typ="INTEGER", value_type="boolean"),
    ],
)

GC_RESULTS = TableSchema(
    logical_name="GC_RESULTS",
    db_table="GC_results",
    sheet_name="GC_results",
    status_datetime_fields=("date",),
    excel_table_name="GC_results",
    columns=[
        _col("resultIdGC", "resultIdGC", typ="INTEGER", value_type="integer", pk=True),
        _col("TaskId", "TaskId", typ="INTEGER", value_type="integer"),
        _col("sampleCode", "sampleCode", typ="TEXT", value_type="text", required=True),
        _col("date", "date", typ="TEXT", value_type="date"),
        _col("operator", "operator", typ="TEXT", value_type="text"),
        _col("equipment", "equipment", typ="TEXT", value_type="text"),
        _col("equipmentNumber", "equipmentNumber", typ="TEXT", value_type="text"),
        _col("Tfact", "Tfact", typ="REAL", value_type="real"),
        _col("presence", "presence", typ="INTEGER", value_type="boolean"),
        _col("deviationABS", "deviationABS", typ="REAL", value_type="real"),
        _col("gasvolume", "gasvolume", typ="REAL", value_type="real"),
        _col("deviationPCT", "deviationPCT", typ="REAL", value_type="real"),
        _col("gasvolumePCT", "gasvolumePCT", typ="REAL", value_type="real"),
        _col("ka1", "ka1", typ="REAL", value_type="real"),
        _col("kb1", "kb1", typ="REAL", value_type="real"),
        _col("ka2", "ka2", typ="REAL", value_type="real"),
        _col("kb2", "kb2", typ="REAL", value_type="real"),
        _col("dateTimeSync", "dateTimeSync", typ="TEXT", value_type="datetime"),
    ],
)

GC_SOURCE_DATA = TableSchema(
    logical_name="GC_SOURCE_DATA",
    db_table="GC_sourceData",
    sheet_name="GC_sourceData",
    parent="GC_RESULTS",
    fk_column="resultIdGC",
    excel_table_name="GC_sourceData",
    delete_if_empty_columns=("volumeH2O", "pressure"),
    columns=[
        _col("rowIdGC", "rowIdGC", typ="INTEGER", value_type="integer", pk=True),
        _col("resultIdGC", "resultIdGC", typ="INTEGER", value_type="integer", required=True),
        _col("volumeH2O", "volumeH2O", typ="REAL", value_type="real"),
        _col("pressure", "pressure", typ="REAL", value_type="real"),
        _col("phase", "phase", typ="TEXT", value_type="text"),
        _col("active", "active", typ="INTEGER", value_type="boolean"),
    ],
)

SSF_RESULTS = TableSchema(
    logical_name="SSF_RESULTS",
    db_table="SSF_results",
    sheet_name="SSF_results",
    status_datetime_fields=("dateStart", "timeStart"),
    excel_table_name="SSF_results",
    columns=[
        _col("resultIdSSF", "resultIdSSF", typ="INTEGER", value_type="integer", pk=True),
        _col("TaskId", "TaskId", typ="INTEGER", value_type="integer"),
        _col("sampleCode", "sampleCode", typ="TEXT", value_type="text", required=True),
        _col("Correct", "Correct", typ="INTEGER", value_type="boolean"),
        _col("dateStart", "dateStart", typ="TEXT", value_type="date"),
        _col("timeStart", "timeStart", typ="TEXT", value_type="time"),
        _col("dateEnd", "dateEnd", typ="TEXT", value_type="date"),
        _col("timeEnd", "timeEnd", typ="TEXT", value_type="time"),
        _col("operator", "operator", typ="TEXT", value_type="text"),
        _col("analyseNumber", "analyseNumber", typ="TEXT", value_type="text"),
        _col("equipmentForTransfer", "equipmentForTransfer", typ="TEXT", value_type="text"),
        _col("equipForTransferNumber", "equipForTransferNumber", typ="TEXT", value_type="text"),
        _col("Ppic", "Ppic", typ="TEXT", value_type="text"),
        _col("PtransferMPa", "PtransferMPa", typ="REAL", value_type="real"),
        _col("Ptransfer", "Ptransfer", typ="REAL", value_type="real"),
        _col("Tfact", "Tfact", typ="REAL", value_type="real"),
        _col("V1", "V1", typ="REAL", value_type="real"),
        _col("V2", "V2", typ="REAL", value_type="real"),
        _col("ssfEquipment", "ssfEquipment", typ="TEXT", value_type="text"),
        _col("ssfEquipmentNumber", "ssfEquipmentNumber", typ="TEXT", value_type="text"),
        _col("Patm", "Patm", typ="REAL", value_type="real"),
        _col("T", "T", typ="REAL", value_type="real"),
        _col("circ2", "circ2", typ="REAL", value_type="real"),
        _col("VHeTot", "VHeTot", typ="REAL", value_type="real"),
        _col("VHeCyl", "VHeCyl", typ="REAL", value_type="real"),
        _col("VGTot", "VGTot", typ="REAL", value_type="real"),
        _col("VGCyl", "VGCyl", typ="REAL", value_type="real"),
        _col("m0Trap1", "m0Trap1", typ="REAL", value_type="real"),
        _col("m1Trap1", "m1Trap1", typ="REAL", value_type="real"),
        _col("m0Trap2", "m0Trap2", typ="REAL", value_type="real"),
        _col("m1Trap2", "m1Trap2", typ="REAL", value_type="real"),
        _col("d20", "d20", typ="REAL", value_type="real"),
        _col("vial", "vial", typ="TEXT", value_type="text"),
        _col("bottle", "bottle", typ="TEXT", value_type="text"),
        _col("pics", "pics", typ="TEXT", value_type="text"),
        _col("Vtransfer", "Vtransfer", typ="REAL", value_type="real"),
        _col("Vgas", "Vgas", typ="REAL", value_type="real"),
        _col("VgasSt", "VgasSt", typ="REAL", value_type="real"),
        _col("mLiq", "mLiq", typ="REAL", value_type="real"),
        _col("VLiq", "VLiq", typ="REAL", value_type="real"),
        _col("GF", "GF", typ="REAL", value_type="real"),
        _col("deltaMatBalance", "deltaMatBalance", typ="REAL", value_type="real"),
        _col("AP", "AP", typ="REAL", value_type="real"),
        _col("dateTimeSync", "dateTimeSync", typ="TEXT", value_type="datetime"),
    ],
)

SSF_SOURCE_DATA = TableSchema(
    logical_name="SSF_SOURCE_DATA",
    db_table="SSF_sourceData",
    sheet_name="SSF_sourceData",
    parent="SSF_RESULTS",
    fk_column="resultIdSSF",
    excel_table_name="SSF_sourceData",
    delete_if_empty_columns=("sampleCode",),
    columns=[
        _col("rowIdSSF", "rowIdSSF", typ="INTEGER", value_type="integer", pk=True),
        _col("resultIdSSF", "resultIdSSF", typ="INTEGER", value_type="integer", required=True),
        _col("sampleCode", "sampleCode", typ="TEXT", value_type="text"),
        _col("TypePh", "TypePh", typ="TEXT", value_type="text"),
    ],
)

GOR_RESULTS = TableSchema(
    logical_name="GOR_RESULTS",
    db_table="GOR_results",
    sheet_name="GOR_results",
    status_datetime_fields=("dateStart", "timeStart"),
    excel_table_name="GOR_results",
    columns=[
        _col("resultIdGOR", "resultIdGOR", typ="INTEGER", value_type="integer", pk=True),
        _col("TaskId", "TaskId", typ="INTEGER", value_type="integer"),
        _col("sampleCode", "sampleCode", typ="TEXT", value_type="text", required=True),
        _col("Correct", "Correct", typ="INTEGER", value_type="boolean"),
        _col("dateStart", "dateStart", typ="TEXT", value_type="date"),
        _col("timeStart", "timeStart", typ="TEXT", value_type="time"),
        _col("dateEnd", "dateEnd", typ="TEXT", value_type="date"),
        _col("timeEnd", "timeEnd", typ="TEXT", value_type="time"),
        _col("operator", "operator", typ="TEXT", value_type="text"),
        _col("analyseNumber", "analyseNumber", typ="TEXT", value_type="text"),
        _col("smplTransfEq", "smplTransfEq", typ="TEXT", value_type="text"),
        _col("transfEqNumber", "transfEqNumber", typ="TEXT", value_type="text"),
        _col("Upic", "Upic", typ="TEXT", value_type="text"),
        _col("Ppic", "Ppic", typ="TEXT", value_type="text"),
        _col("PtransferMPa", "PtransferMPa", typ="REAL", value_type="real"),
        _col("Ptransfer", "Ptransfer", typ="REAL", value_type="real"),
        _col("Tfact", "Tfact", typ="REAL", value_type="real"),
        _col("V1", "V1", typ="REAL", value_type="real"),
        _col("V2", "V2", typ="REAL", value_type="real"),
        _col("m0", "m0", typ="REAL", value_type="real"),
        _col("m1", "m1", typ="REAL", value_type="real"),
        _col("gorEquipment", "gorEquipment", typ="TEXT", value_type="text"),
        _col("gorEquipmentNumber", "gorEquipmentNumber", typ="TEXT", value_type="text"),
        _col("Patm", "Patm", typ="REAL", value_type="real"),
        _col("T", "T", typ="REAL", value_type="real"),
        _col("circ1", "circ1", typ="REAL", value_type="real"),
        _col("circ2", "circ2", typ="REAL", value_type="real"),
        _col("VHeTot", "VHeTot", typ="REAL", value_type="real"),
        _col("VHeCyl", "VHeCyl", typ="REAL", value_type="real"),
        _col("VGTot", "VGTot", typ="REAL", value_type="real"),
        _col("VGCyl", "VGCyl", typ="REAL", value_type="real"),
        _col("m0Trap", "m0Trap", typ="REAL", value_type="real"),
        _col("m1Trap", "m1Trap", typ="REAL", value_type="real"),
        _col("m2", "m2", typ="REAL", value_type="real"),
        _col("d20", "d20", typ="REAL", value_type="real"),
        _col("vial", "vial", typ="TEXT", value_type="text"),
        _col("bottle", "bottle", typ="TEXT", value_type="text"),
        _col("TedlarBag", "TedlarBag", typ="TEXT", value_type="text"),
        _col("pics", "pics", typ="TEXT", value_type="text"),
        _col("Vtransfer", "Vtransfer", typ="REAL", value_type="real"),
        _col("Vgas", "Vgas", typ="REAL", value_type="real"),
        _col("VgasSt", "VgasSt", typ="REAL", value_type="real"),
        _col("mgas", "mgas", typ="REAL", value_type="real"),
        _col("dGas", "dGas", typ="REAL", value_type="real"),
        _col("VLiq", "VLiq", typ="REAL", value_type="real"),
        _col("mLiq", "mLiq", typ="REAL", value_type="real"),
        _col("GF", "GF", typ="REAL", value_type="real"),
        _col("deltaMatBalance", "deltaMatBalance", typ="REAL", value_type="real"),
        _col("deltaAP", "deltaAP", typ="REAL", value_type="real"),
        _col("AP", "AP", typ="REAL", value_type="real"),
        _col("picDen", "picDen", typ="REAL", value_type="real"),
        _col("KGF", "KGF", typ="REAL", value_type="real"),
        _col("dateTimeSync", "dateTimeSync", typ="TEXT", value_type="datetime"),
    ],
)

GOR_SOURCE_DATA = TableSchema(
    logical_name="GOR_SOURCE_DATA",
    db_table="GOR_sourceData",
    sheet_name="GOR_sourceData",
    parent="GOR_RESULTS",
    fk_column="resultIdGOR",
    excel_table_name="GOR_sourceData",
    delete_if_empty_columns=("sampleCode",),
    columns=[
        _col("rowIdGOR", "rowIdGOR", typ="INTEGER", value_type="integer", pk=True),
        _col("resultIdGOR", "resultIdGOR", typ="INTEGER", value_type="integer", required=True),
        _col("sampleCode", "sampleCode", typ="TEXT", value_type="text"),
        _col("TypePh", "TypePh", typ="TEXT", value_type="text"),
    ],
)

BP_RESULTS = TableSchema(
    logical_name="BP_RESULTS",
    db_table="BP_results",
    sheet_name="BP_results",
    status_datetime_fields=("date",),
    excel_table_name="BP_results",
    columns=[
        _col("resultIdBP", "resultIdBP", typ="INTEGER", value_type="integer", pk=True),
        _col("TaskId", "TaskId", typ="INTEGER", value_type="integer"),
        _col("sampleCode", "sampleCode", typ="TEXT", value_type="text", required=True),
        _col("date", "date", typ="TEXT", value_type="date"),
        _col("operator", "operator", typ="TEXT", value_type="text"),
        _col("analyseNumber", "analyseNumber", typ="TEXT", value_type="text"),
        _col("bpEquipment", "bpEquipment", typ="TEXT", value_type="text"),
        _col("bpEquipmentNumber", "bpEquipmentNumber", typ="TEXT", value_type="text"),
        _col("Tfact", "Tfact", typ="REAL", value_type="real"),
        _col("bpMPa", "bpMPa", typ="REAL", value_type="real"),
        _col("Correct", "Correct", typ="INTEGER", value_type="boolean"),
        _col("deltaPCT", "deltaPCT", typ="REAL", value_type="real"),
        _col("ka1", "ka1", typ="REAL", value_type="real"),
        _col("kb1", "kb1", typ="REAL", value_type="real"),
        _col("ka2", "ka2", typ="REAL", value_type="real"),
        _col("kb2", "kb2", typ="REAL", value_type="real"),
        _col("dateTimeSync", "dateTimeSync", typ="TEXT", value_type="datetime"),
    ],
)

BP_SOURCE_DATA = TableSchema(
    logical_name="BP_SOURCE_DATA",
    db_table="BP_sourceData",
    sheet_name="BP_sourceData",
    parent="BP_RESULTS",
    fk_column="resultIdBP",
    excel_table_name="BP_sourceData",
    delete_if_empty_columns=("pressure", "heightPiston"),
    columns=[
        _col("rowIdBP", "rowIdBP", typ="INTEGER", value_type="integer", pk=True),
        _col("resultIdBP", "resultIdBP", typ="INTEGER", value_type="integer", required=True),
        _col("pressure", "pressure", typ="REAL", value_type="real"),
        _col("heightPiston", "heightPiston", typ="REAL", value_type="real"),
        _col("phase", "phase", typ="TEXT", value_type="text"),
        _col("active", "active", typ="INTEGER", value_type="boolean"),
    ],
)

REC_RESULTS = TableSchema(
    logical_name="REC_RESULTS",
    db_table="Rec_results",
    sheet_name="Rec_results",
    excel_table_name="REC_results",
    status_datetime_fields=("date",),
    columns=[
        _col("resultIdRec", "resultIdRec", typ="INTEGER", value_type="integer", pk=True),
        _col("TaskId", "TaskId", typ="INTEGER", value_type="integer"),
        _col("sampleCode", "sampleCode", typ="TEXT", value_type="text", required=True),
        _col("operator", "operator", typ="TEXT", value_type="text"),
        _col("date", "date", typ="TEXT", value_type="date"),
        _col("equipmentForTransfer", "equipmentForTransfer", typ="TEXT", value_type="text"),
        _col("equipmentNumber", "equipmentNumber", typ="TEXT", value_type="text"),
        _col("typeSampler", "typeSampler", typ="TEXT", value_type="text"),
        _col("numberSampler", "numberSampler", typ="TEXT", value_type="text"),
        _col("dateTimeSync", "dateTimeSync", typ="TEXT", value_type="datetime"),
    ],
)

REC_SOURCE_DATA = TableSchema(
    logical_name="REC_SOURCE_DATA",
    db_table="Rec_sourceData",
    sheet_name="Rec_sourceData",
    parent="REC_RESULTS",
    fk_column="resultIdRec",
    excel_table_name="REC_sourceData",
    delete_if_empty_columns=("mixSampleCode", "mTransfered" ),
    columns=[
        _col("rowIdRec", "rowIdRec", typ="INTEGER", value_type="integer", pk=True),
        _col("resultIdRec", "resultIdRec", typ="INTEGER", value_type="integer", required=True),
        _col("massFraction", "massFraction", typ="REAL", value_type="real"),
        _col("mixSampleCode", "mixSampleCode", typ="TEXT", value_type="text"),
        _col("Ptransfer", "Ptransfer", typ="REAL", value_type="real"),
        _col("Ttransfer", "Ttransfer", typ="REAL", value_type="real"),
        _col("Density", "Density", typ="REAL", value_type="real"),
        _col("first", "first", typ="INTEGER", value_type="boolean"),
        _col("Volume", "Volume", typ="REAL", value_type="real"),
        _col("mass", "mass", typ="REAL", value_type="real"),
    ],
)

EMV_RESULTS = TableSchema(
    logical_name="EMV_RESULTS",
    db_table="EMV_results",
    sheet_name="EMV_results",
    excel_table_name="EMV_results",
    status_datetime_fields=("date",),
    columns=[
        _col("resultIdEMV", "resultIdEMV", typ="INTEGER", value_type="integer", pk=True),
        _col("TaskId", "TaskId", typ="INTEGER", value_type="integer"),
        _col("sampleCode", "sampleCode", typ="TEXT", value_type="text", required=True),
        _col("date", "date", typ="TEXT", value_type="date"),
        _col("operator", "operator", typ="TEXT", value_type="text"),
        _col("analyseNumber", "analyseNumber", typ="TEXT", value_type="text"),
        _col("emvEquipment", "emvEquipment", typ="TEXT", value_type="text"),
        _col("emvEquipmentNumber", "emvEquipmentNumber", typ="TEXT", value_type="text"),
        _col("Tfact", "Tfact", typ="REAL", value_type="real"),
        _col("Correct", "Correct", typ="INTEGER", value_type="boolean"),
        _col("ka", "ka", typ="REAL", value_type="real"),
        _col("kb", "kb", typ="REAL", value_type="real"),
        _col("dateTimeSync", "dateTimeSync", typ="TEXT", value_type="datetime"),
    ],
)

EMV_SOURCE_DATA = TableSchema(
    logical_name="EMV_SOURCE_DATA",
    db_table="EMV_sourceData",
    sheet_name="EMV_sourceData",
    parent="EMV_RESULTS",
    fk_column="resultIdEMV",
    excel_table_name="EMV_sourceData",
    delete_if_empty_columns=("pressure", "viscosity"),
    columns=[
        _col("rowIdEMV", "rowIdEMV", typ="INTEGER", value_type="integer", pk=True),
        _col("resultIdEMV", "resultIdEMV", typ="INTEGER", value_type="integer", required=True),
        _col("pressure", "pressure", typ="REAL", value_type="real"),
        _col("temperature", "temperature", typ="REAL", value_type="real"),
        _col("viscosity", "viscosity", typ="REAL", value_type="real"),
        _col("active", "active", typ="INTEGER", value_type="boolean"),
    ],
)

JOIN_RESULTS = TableSchema(
logical_name="JOIN_RESULTS",
    db_table="JOIN_results",
    sheet_name="JOIN_results",
    excel_table_name="JOIN_results",
    status_datetime_fields=("date",),
    columns=[
        _col("resultIdJOIN", "resultIdJOIN", typ="INTEGER", value_type="integer", pk=True),
        _col("TaskId", "TaskId", typ="INTEGER", value_type="integer"),
        _col("sampleCode", "sampleCode", typ="TEXT", value_type="text", required=True),
        _col("date", "date", typ="TEXT", value_type="date"),
        _col("operator", "operator", typ="TEXT", value_type="text"),
        _col("equipment", "equipment", typ="TEXT", value_type="text"),
        _col("equipmentNumber", "equipmentNumber", typ="TEXT", value_type="text"),
        _col("preparationPressure", "preparationPressure", typ="REAL", value_type="real"),
        _col("absolutePressure", "absolutePressure", typ="INTEGER", value_type="boolean"),
        _col("pressureUnit", "pressureUnit", typ="TEXT", value_type="text"),
        _col("preparationTemperature", "preparationTemperature", typ="REAL", value_type="real"),
        _col("preparationPressureMPa", "preparationPressureMPa", typ="REAL", value_type="real"),
        _col("completeFWC", "completeFWC", typ="INTEGER", value_type="boolean"),
        _col("dateTimeSync", "dateTimeSync", typ="TEXT", value_type="datetime"),
    ],
)

JOIN_SOURCE_DATA = TableSchema(
    logical_name="JOIN_SOURCE_DATA",
    db_table="JOIN_sourceData",
    sheet_name="JOIN_sourceData",
    parent="JOIN_RESULTS",
    fk_column="resultIdJOIN",
    excel_table_name="JOIN_sourceData",
    delete_if_empty_columns=("mixSampleCode",),
    columns=[
        _col("rowIdJOIN", "rowIdJOIN", typ="INTEGER", value_type="integer", pk=True),
        _col("resultIdJOIN", "resultIdJOIN", typ="INTEGER", value_type="integer", required=True),
        _col("mixSampleCode", "mixSampleCode", typ="TEXT", value_type="text"),
        _col("depleted", "depleted", typ="INTEGER", value_type="boolean"),
    ],
)

CCE_RESULTS = TableSchema(
    logical_name="CCE_RESULTS",
    db_table="CCE_results",
    sheet_name="CCE_results",
    excel_table_name="CCE_results",
    status_datetime_fields=("dateEnd",),
    columns=[
        _col(db_name="resultIdCCE", excel_name="resultIdCCE", typ="INTEGER", value_type="integer", pk=True),
        _col(db_name="TaskId", excel_name="TaskId", typ="INTEGER", value_type="integer"),
        _col(db_name="sampleCode", excel_name="sampleCode", typ="TEXT", value_type="text", required=True),
        _col(db_name="PVTEqNumber", excel_name="PVTEqNumber", typ="TEXT", value_type="text"),

        _col(db_name="dateStart", excel_name="dateStart", typ="TEXT", value_type="date"),
        _col(db_name="dateEnd", excel_name="dateEnd", typ="TEXT", value_type="date"),
        _col(db_name="operator", excel_name="operator", typ="TEXT", value_type="text"),
        _col(db_name="analyseNumber", excel_name="analyseNumber", typ="TEXT", value_type="text"),
        _col(db_name="typeSample", excel_name="typeSample", typ="TEXT", value_type="text"),

        _col(db_name="Vres", excel_name="Vres", typ="REAL", value_type="real"),
        _col(db_name="Vdew", excel_name="Vdew", typ="REAL", value_type="real"),

        _col(db_name="Pres", excel_name="Pres", typ="REAL", value_type="real"),
        _col(db_name="unitPressure", excel_name="unitPressure", typ="TEXT", value_type="text"),
        _col(db_name="PresMPaABS", excel_name="PresMPaABS", typ="REAL", value_type="real"),
        _col(db_name="Tfact", excel_name="Tfact", typ="REAL", value_type="real"),

        _col(db_name="fluidMass", excel_name="fluidMass", typ="REAL", value_type="real"),
        _col(db_name="fluidTransferDensity", excel_name="fluidTransferDensity", typ="REAL", value_type="real"),

        _col(db_name="dateTimeSync", excel_name="dateTimeSync", typ="TEXT", value_type="datetime"),
    ],
)

CCE_SOURCE_DATA = TableSchema(
    logical_name="CCE_SOURCE_DATA",
    db_table="CCE_sourceData",
    sheet_name="CCE_sourceData",
    parent="CCE_RESULTS",
    fk_column="resultIdCCE",
    excel_table_name="CCE_sourceData",
    delete_if_empty_columns=("Pressure", "Vcell"),
    columns=[
        _col(db_name="rowIdCCE", excel_name="rowIdCCE", typ="INTEGER", value_type="integer", pk=True),
        _col(db_name="resultIdCCE", excel_name="resultIdCCE", typ="INTEGER", value_type="integer", required=True),

        _col(db_name="Pressure", excel_name="Pressure", typ="REAL", value_type="real"),
        _col(db_name="PressureMPa", excel_name="PressureMPa", typ="REAL", value_type="real"),
        _col(db_name="Phase", excel_name="Phase", typ="INTEGER", value_type="integer"),

        _col(db_name="Vcell", excel_name="Vcell", typ="REAL", value_type="real"),
        _col(db_name="Vliquid", excel_name="Vliquid", typ="REAL", value_type="real"),
        _col(db_name="Vgas", excel_name="Vgas", typ="REAL", value_type="real"),

        _col(db_name="relativeCellVolume", excel_name="relativeCellVolume", typ="REAL", value_type="real"),
        _col(db_name="liquidVolumeFraction", excel_name="liquidVolumeFraction", typ="REAL", value_type="real"),
        _col(db_name="gasVolumeFraction", excel_name="gasVolumeFraction", typ="REAL", value_type="real"),

        _col(db_name="monophaseDensity", excel_name="monophaseDensity", typ="REAL", value_type="real"),
        _col(db_name="comment", excel_name="comment", typ="TEXT", value_type="text"),
        _col(db_name="active", excel_name="active", typ="INTEGER", value_type="boolean"),
    ],
)

CVD_RESULTS = TableSchema(
    logical_name="CVD_RESULTS",
    db_table="CVD_results",
    sheet_name="CVD_results",
    excel_table_name="CVD_results",
    status_datetime_fields=("dateStart", "timeStart"),
    columns=[
        _col("resultIdCVD", "resultIdCVD", typ="INTEGER", value_type="integer", pk=True),

        _col("sampleCode", "sampleCode", typ="TEXT", value_type="text", required=True),
        _col("TaskId", "TaskId", typ="INTEGER", value_type="integer"),
        _col("analyseNumber", "analyseNumber", typ="INTEGER", value_type="integer", required=True),

        _col("dateEnd", "dateEnd", typ="TEXT", value_type="date"),
        _col("dateStart", "dateStart", typ="TEXT", value_type="date"),
        _col("timeEnd", "timeEnd", typ="TEXT", value_type="time"),
        _col("timeStart", "timeStart", typ="TEXT", value_type="time"),

        _col("deltaMatBalance", "deltaMatBalance", typ="REAL", value_type="real"),

        _col("Pabs", "Pabs", typ="INTEGER", value_type="boolean"),
        _col("Pstart", "Pstart", typ="REAL", value_type="real"),
        _col("Ptransfer", "Ptransfer", typ="REAL", value_type="real"),
        _col("Punit", "Punit", typ="TEXT", value_type="text"),

        _col("pvtCell", "pvtCell", typ="TEXT", value_type="text"),

        _col("Tstart", "Tstart", typ="REAL", value_type="real"),
        _col("Ttransfer", "Ttransfer", typ="REAL", value_type="real"),

        _col("Vstart", "Vstart", typ="REAL", value_type="real"),
        _col("Vtransfer", "Vtransfer", typ="REAL", value_type="real"),

        _col("dateTimeSync", "dateTimeSync", typ="TEXT", value_type="datetime"),
    ],
)

CVD_SOURCE_DATA = TableSchema(
    logical_name="CVD_SOURCE_DATA",
    db_table="CVD_sourceData",
    sheet_name="CVD_sourceData",
    excel_table_name="CVD_sourceData",
    parent="CVD_RESULTS",
    fk_column="resultIdCVD",

    # Важно: step не включаю сюда специально.
    # Если step будет заранее заполнен во всех строках, иначе пустые строки будут считаться "данными".
    delete_if_empty_columns=(
        "d20",
        "date",
        "gorEquipment",
        "initPatm",
        "m0",
        "m0Trap",
        "m1",
        "m1Trap",
        "m2",
        "operator",
        "Patm",
        "sampleCodeGPh",
        "sampleCodeLPh",
        "stepPcell",
        "stepTcell",
        "T",
        "transferPcell",
        "V1",
        "V2",
        "Vcell",
        "Vcellend",
        "VCellLiq",
        "VGCyl",
        "VGTot",
    ),
    columns=[
        _col("rowIdCVD", "rowIdCVD", typ="INTEGER", value_type="integer", pk=True),

        _col("resultIdCVD", "resultIdCVD", typ="INTEGER", value_type="integer", required=True),

        _col("step", "step", typ="TEXT", value_type="text"),

        _col("cellFluidMass", "cellFluidMass", typ="REAL", value_type="real"),
        _col("d20", "d20", typ="REAL", value_type="real"),
        _col("date", "date", typ="TEXT", value_type="date"),
        _col("dGas", "dGas", typ="REAL", value_type="real"),
        _col("evacuatedFluidMass", "evacuatedFluidMass", typ="REAL", value_type="real"),

        _col("GF", "GF", typ="REAL", value_type="real"),
        _col("gorEquipment", "gorEquipment", typ="TEXT", value_type="text"),

        _col("initPatm", "initPatm", typ="REAL", value_type="real"),

        _col("m0", "m0", typ="REAL", value_type="real"),
        _col("m0Trap", "m0Trap", typ="REAL", value_type="real"),
        _col("m1", "m1", typ="REAL", value_type="real"),
        _col("m1Trap", "m1Trap", typ="REAL", value_type="real"),
        _col("m2", "m2", typ="REAL", value_type="real"),

        _col("mgas", "mgas", typ="REAL", value_type="real"),
        _col("mLiq", "mLiq", typ="REAL", value_type="real"),

        _col("operator", "operator", typ="TEXT", value_type="text"),

        _col("Patm", "Patm", typ="REAL", value_type="real"),
        _col("picDen", "picDen", typ="REAL", value_type="real"),
        _col("Ppic", "Ppic", typ="TEXT", value_type="text"),

        _col("sampleCodeGPh", "sampleCodeGPh", typ="TEXT", value_type="text"),
        _col("sampleCodeLPh", "sampleCodeLPh", typ="TEXT", value_type="text"),

        _col("stepPcell", "stepPcell", typ="REAL", value_type="real"),
        _col("stepTcell", "stepTcell", typ="REAL", value_type="real"),

        _col("T", "T", typ="REAL", value_type="real"),
        _col("transferPcell", "transferPcell", typ="REAL", value_type="real"),

        _col("Upic", "Upic", typ="TEXT", value_type="text"),

        _col("V1", "V1", typ="REAL", value_type="real"),
        _col("V2", "V2", typ="REAL", value_type="real"),
        _col("Vcell", "Vcell", typ="REAL", value_type="real"),
        _col("Vcellend", "Vcellend", typ="REAL", value_type="real"),
        _col("VCellLiq", "VCellLiq", typ="REAL", value_type="real"),

        _col("Vgas", "Vgas", typ="REAL", value_type="real"),
        _col("VgasSt", "VgasSt", typ="REAL", value_type="real"),

        _col("VGCyl", "VGCyl", typ="REAL", value_type="real"),
        _col("VGTot", "VGTot", typ="REAL", value_type="real"),

        _col("VHeCyl", "VHeCyl", typ="REAL", value_type="real"),
        _col("VHeTot", "VHeTot", typ="REAL", value_type="real"),

        _col("VLiq", "VLiq", typ="REAL", value_type="real"),
        _col("VLiqRelative", "VLiqRelative", typ="REAL", value_type="real"),

        _col("Vtransfer", "Vtransfer", typ="REAL", value_type="real"),
    ],
)


STUDIES: dict[str, StudySchema] = {
    "OP": StudySchema("OP", OP_RESULTS, None),
    "AP": StudySchema("AP", AP_RESULTS, AP_SOURCE_DATA),
    "GC": StudySchema("GC", GC_RESULTS, GC_SOURCE_DATA),
    "SSF": StudySchema("SSF", SSF_RESULTS, SSF_SOURCE_DATA),
    "GOR": StudySchema("GOR", GOR_RESULTS, GOR_SOURCE_DATA),
    "BP": StudySchema("BP", BP_RESULTS, BP_SOURCE_DATA),
    "REC": StudySchema("REC", REC_RESULTS, REC_SOURCE_DATA),
    "EMV": StudySchema("EMV", EMV_RESULTS, EMV_SOURCE_DATA),
    "JOIN": StudySchema("JOIN", JOIN_RESULTS, JOIN_SOURCE_DATA),
    "CCE": StudySchema("CCE", CCE_RESULTS, CCE_SOURCE_DATA),
    "OPOH": StudySchema("OPOH", OPOH_RESULTS, None),
    "CVD": StudySchema(
        code="CVD",
        result_table=CVD_RESULTS,
        source_table=CVD_SOURCE_DATA,
        progress_step_field="step",
        progress_total_steps=10,
        completion_step=11,
    ),
}

def get_studies(codes: list[str] | None = None) -> list[StudySchema]:
    """
    Возвращает исследования по кодам.

    Если codes=None или список пустой, возвращаются все исследования из STUDIES.
    """
    if not codes:
        return list(STUDIES.values())

    result: list[StudySchema] = []
    unknown: list[str] = []

    for code in codes:
        normalized = code.strip().upper()
        study = STUDIES.get(normalized)
        if study is None:
            unknown.append(code)
        else:
            result.append(study)

    if unknown:
        raise KeyError(f"Неизвестные типы исследований: {', '.join(unknown)}")

    return result
