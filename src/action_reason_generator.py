from pathlib import Path

from src.dat_generator import generate_hdl_bundle
from src.field_mapper import normalize_column_name
from src.validator import format_hdl_date, is_empty


ACTION_REASON_FIELDS = [
    "SourceSystemOwner",
    "SourceSystemId",
    "ActionReasonId",
    "ActionReasonCode",
    "ActionReason",
    "StartDate",
    "EndDate",
]

_ACTION_REASON_SOURCE_ALIASES = {
    "ActionReasonCode": ("ActionReasonCode",),
    "ActionReason": ("ActionReason",),
    "StartDate": ("ActionReasonStartDate", "StartDate"),
    "EndDate": ("ActionReasonEndDate", "EndDate"),
}


def build_action_reason_output_objects(dataframe, source_system_owner):
    """Build distinct ActionReasons from workbook action-reason columns."""
    if not source_system_owner:
        raise ValueError("A source system owner is required.")

    normalized_columns = {
        normalize_column_name(column): column
        for column in dataframe.columns
    }
    source_columns = {
        field: next(
            (
                normalized_columns[normalize_column_name(alias)]
                for alias in aliases
                if normalize_column_name(alias) in normalized_columns
            ),
            None,
        )
        for field, aliases in _ACTION_REASON_SOURCE_ALIASES.items()
    }
    missing_columns = [
        field for field, column in source_columns.items() if column is None
    ]
    if missing_columns:
        raise ValueError(
            "ActionReasons requires Excel columns: "
            + ", ".join(missing_columns)
        )

    records = []
    records_by_code = {}
    for row_index, (_, row) in enumerate(dataframe.iterrows(), start=2):
        source_values = {
            field: row[column]
            for field, column in source_columns.items()
        }
        if all(is_empty(value) for value in source_values.values()):
            continue

        missing_values = [
            field for field, value in source_values.items()
            if is_empty(value)
        ]
        if missing_values:
            raise ValueError(
                f"Excel row {row_index}: ActionReasons fields must not be "
                "blank: " + ", ".join(missing_values)
            )

        code = str(source_values["ActionReasonCode"]).strip()
        reason = str(source_values["ActionReason"]).strip()
        start_date = source_values["StartDate"]
        end_date = source_values["EndDate"]
        code_key = code.casefold()
        signature = (
            reason,
            format_hdl_date(start_date),
            format_hdl_date(end_date),
        )
        existing = records_by_code.get(code_key)
        if existing is not None:
            if existing["signature"] != signature:
                raise ValueError(
                    f"Excel row {row_index}: conflicting ActionReasons data "
                    f"for ActionReasonCode '{code}'."
                )
            continue

        source_system_id = f"{source_system_owner}_AR_{100 + len(records) + 1:03d}"
        record = {
            "SourceSystemOwner": source_system_owner,
            "SourceSystemId": source_system_id,
            "ActionReasonId": "",
            "ActionReasonCode": code,
            "ActionReason": reason,
            "StartDate": start_date,
            "EndDate": end_date,
        }
        records_by_code[code_key] = {
            "signature": signature,
            "record": record,
        }
        records.append(record)

    return [{
        "object_name": "ActionReasons",
        "metadata_fields": ACTION_REASON_FIELDS,
        "records": records,
    }]


def generate_action_reasons_dat(dataframe, source_system_owner, output_dir):
    """Write ActionReasons.dat from workbook action-reason columns."""
    output_file = Path(output_dir) / "ActionReasons.dat"
    return generate_hdl_bundle(
        output_file,
        build_action_reason_output_objects(dataframe, source_system_owner),
    )
