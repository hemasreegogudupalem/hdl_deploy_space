import re
from pathlib import Path

from src.dat_generator import generate_hdl_bundle
from src.field_mapper import normalize_column_name
from src.validator import format_hdl_date, is_empty


ACTIONS_FIELDS = [
    "FLEX:PER_ACT_LEG_DDF",
    "ActionId",
    "ActionCode",
    "ActionName",
    "Description",
    "ActionTypeId",
    "ActionTypeCode",
    "StartDate",
    "EndDate",
    "GUID",
    "SourceSystemId",
    "SourceSystemOwner",
]

_ACTIONS_SOURCE_ALIASES = {
    "ActionCode": ("ActionCode",),
    "ActionName": ("ActionName",),
    "Description": ("ActionDescription", "Description"),
    "ActionTypeCode": ("ActionTypeCode",),
    "StartDate": ("ActionStartDate", "StartDate"),
    "EndDate": ("ActionEndDate", "EndDate"),
}


def build_actions_output_objects(dataframe, source_system_owner):
    """Build distinct Actions records from the selected workbook."""
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
        for field, aliases in _ACTIONS_SOURCE_ALIASES.items()
    }
    missing_columns = [
        field for field, column in source_columns.items() if column is None
    ]
    if missing_columns:
        raise ValueError(
            "Actions requires Excel columns: " + ", ".join(missing_columns)
        )

    action_type_id_column = normalized_columns.get(
        normalize_column_name("ActionTypeId")
    )
    records = []
    records_by_code = {}
    seen_source_ids = set()

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
                f"Excel row {row_index}: Actions fields must not be blank: "
                + ", ".join(missing_values)
            )

        action_code = str(source_values["ActionCode"]).strip()
        action_name = str(source_values["ActionName"]).strip()
        description = str(source_values["Description"]).strip()
        action_type_code = str(source_values["ActionTypeCode"]).strip()
        start_date = source_values["StartDate"]
        end_date = source_values["EndDate"]
        action_type_id = ""
        if action_type_id_column is not None:
            value = row[action_type_id_column]
            if not is_empty(value):
                action_type_id = str(value).strip()
                if action_type_id.endswith(".0"):
                    action_type_id = action_type_id[:-2]

        action_key = action_code.casefold()
        signature = (
            action_name,
            description,
            action_type_id,
            action_type_code,
            format_hdl_date(start_date),
            format_hdl_date(end_date),
        )
        existing = records_by_code.get(action_key)
        if existing is not None:
            if existing["signature"] != signature:
                raise ValueError(
                    f"Excel row {row_index}: conflicting Actions data for "
                    f"ActionCode '{action_code}'."
                )
            continue

        safe_code = re.sub(
            r"[^A-Z0-9]+",
            "_",
            action_code.upper(),
        ).strip("_")
        if not safe_code:
            raise ValueError(
                f"Excel row {row_index}: ActionCode cannot be used to "
                "generate a SourceSystemId."
            )
        source_system_id = f"ACT_{safe_code}_001"
        if source_system_id in seen_source_ids:
            raise ValueError(
                f"Excel row {row_index}: duplicate SourceSystemId "
                f"'{source_system_id}'."
            )
        seen_source_ids.add(source_system_id)

        record = {
            "ActionId": "",
            "ActionCode": action_code,
            "ActionName": action_name,
            "Description": description,
            "ActionTypeId": action_type_id,
            "ActionTypeCode": action_type_code,
            "StartDate": start_date,
            "EndDate": end_date,
            "GUID": "",
            "SourceSystemId": source_system_id,
            "SourceSystemOwner": source_system_owner,
        }
        records_by_code[action_key] = {
            "signature": signature,
            "record": record,
        }
        records.append(record)

    return [{
        "object_name": "Actions",
        "metadata_fields": ACTIONS_FIELDS,
        "records": records,
    }]


def generate_actions_dat(dataframe, source_system_owner, output_dir):
    """Write Actions.dat from action definitions in workbook columns."""
    output_file = Path(output_dir) / "Actions.dat"
    return generate_hdl_bundle(
        output_file,
        build_actions_output_objects(dataframe, source_system_owner),
    )
