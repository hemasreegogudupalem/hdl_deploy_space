from pathlib import Path

from src.dat_generator import generate_hdl_bundle
from src.field_mapper import normalize_column_name
from src.validator import is_empty


CALENDAR_EVENT_FIELDS = [
    "Name",
    "Description",
    "Category",
    "CoverageType",
    "ShortCode",
    "StartDateTime",
    "EndDateTime",
    "TreeCode",
    "TreeStructureCode",
    "TreeVersionName",
    "HalfDayForElapsed",
]

CALENDAR_EVENT_COVERAGE_FIELDS = [
    "ShortCode",
    "CoverageScope",
    "TerritoryCode",
    "TreeStructureCode",
    "TreeCode",
    "TreeVersionName",
]

_CALENDAR_EVENT_REQUIRED_FIELDS = (
    "Name",
    "Description",
    "Category",
    "CoverageType",
    "ShortCode",
    "StartDateTime",
    "EndDateTime",
    "HalfDayForElapsed",
)

_CALENDAR_EVENT_COVERAGE_REQUIRED_FIELDS = (
    "ShortCode",
    "CoverageScope",
    "TerritoryCode",
)


def _column_map(dataframe):
    return {
        normalize_column_name(column): column
        for column in dataframe.columns
    }


def _read_record(row, fields, columns):
    return {
        field: row[columns[normalize_column_name(field)]]
        if normalize_column_name(field) in columns
        else ""
        for field in fields
    }


def _validate_required_columns(columns, fields, object_name):
    missing_columns = [
        field for field in fields
        if normalize_column_name(field) not in columns
    ]
    if missing_columns:
        raise ValueError(
            f"{object_name} requires Excel columns: "
            + ", ".join(missing_columns)
        )


def build_calendar_event_output_objects(event_dataframe, coverage_dataframe):
    """Build CalendarEvent and coverage records from their workbook sheets."""
    event_columns = _column_map(event_dataframe)
    coverage_columns = _column_map(coverage_dataframe)
    _validate_required_columns(
        event_columns,
        _CALENDAR_EVENT_REQUIRED_FIELDS,
        "CalendarEvent",
    )
    _validate_required_columns(
        coverage_columns,
        _CALENDAR_EVENT_COVERAGE_REQUIRED_FIELDS,
        "CalendarEventCoverage",
    )

    event_records = []
    events_by_code = {}
    for row_index, (_, row) in enumerate(event_dataframe.iterrows(), start=2):
        record = _read_record(row, CALENDAR_EVENT_FIELDS, event_columns)
        required_values = {
            field: record[field]
            for field in _CALENDAR_EVENT_REQUIRED_FIELDS
        }
        if all(is_empty(value) for value in required_values.values()):
            continue

        missing_values = [
            field for field, value in required_values.items()
            if is_empty(value)
        ]
        if missing_values:
            raise ValueError(
                f"Excel row {row_index}: CalendarEvent fields must not be "
                "blank: " + ", ".join(missing_values)
            )

        short_code = str(record["ShortCode"]).strip()
        code_key = short_code.casefold()
        if code_key in events_by_code:
            raise ValueError(
                f"Excel row {row_index}: duplicate CalendarEvent ShortCode "
                f"'{short_code}'."
            )
        record["ShortCode"] = short_code
        for field in CALENDAR_EVENT_FIELDS:
            if field not in {"StartDateTime", "EndDateTime"}:
                record[field] = (
                    str(record[field]).strip()
                    if not is_empty(record[field])
                    else ""
                )
        event_records.append(record)
        events_by_code[code_key] = record

    coverage_records = []
    coverage_codes = set()
    for row_index, (_, row) in enumerate(
        coverage_dataframe.iterrows(),
        start=2,
    ):
        record = _read_record(
            row,
            CALENDAR_EVENT_COVERAGE_FIELDS,
            coverage_columns,
        )
        required_values = {
            field: record[field]
            for field in _CALENDAR_EVENT_COVERAGE_REQUIRED_FIELDS
        }
        if all(is_empty(value) for value in required_values.values()):
            continue

        missing_values = [
            field for field, value in required_values.items()
            if is_empty(value)
        ]
        if missing_values:
            raise ValueError(
                f"Excel row {row_index}: CalendarEventCoverage fields must "
                "not be blank: " + ", ".join(missing_values)
            )

        short_code = str(record["ShortCode"]).strip()
        code_key = short_code.casefold()
        if code_key in coverage_codes:
            raise ValueError(
                f"Excel row {row_index}: duplicate CalendarEventCoverage "
                f"ShortCode '{short_code}'."
            )
        if code_key not in events_by_code:
            raise ValueError(
                f"Excel row {row_index}: no CalendarEvent found for "
                f"ShortCode '{short_code}'."
            )
        record["ShortCode"] = short_code
        for field in CALENDAR_EVENT_COVERAGE_FIELDS:
            if field != "ShortCode":
                record[field] = (
                    str(record[field]).strip()
                    if not is_empty(record[field])
                    else ""
                )
        coverage_records.append(record)
        coverage_codes.add(code_key)

    missing_coverage = [
        record["ShortCode"]
        for record in event_records
        if record["ShortCode"].casefold() not in coverage_codes
    ]
    if missing_coverage:
        raise ValueError(
            "CalendarEventCoverage is missing ShortCode values: "
            + ", ".join(missing_coverage)
        )

    return [
        {
            "object_name": "CalendarEvent",
            "metadata_fields": CALENDAR_EVENT_FIELDS,
            "records": event_records,
        },
        {
            "object_name": "CalendarEventCoverage",
            "metadata_fields": CALENDAR_EVENT_COVERAGE_FIELDS,
            "records": coverage_records,
        },
    ]


def generate_calendar_event_dat(
    event_dataframe,
    coverage_dataframe,
    output_dir,
):
    """Write CalendarEvent.dat with event and coverage HDL records."""
    output_file = Path(output_dir) / "CalendarEvent.dat"
    return generate_hdl_bundle(
        output_file,
        build_calendar_event_output_objects(
            event_dataframe,
            coverage_dataframe,
        ),
    )
