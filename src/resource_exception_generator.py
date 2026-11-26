from pathlib import Path

from src.dat_generator import generate_hdl_bundle
from src.field_mapper import normalize_column_name
from src.validator import is_empty


RESOURCE_EXCEPTION_FIELDS = [
    "ExceptionId",
    "ExceptionName",
    "StartDateTime",
    "EndDateTime",
    "AvailabilityCode",
    "SourceSystemOwner",
    "SourceSystemId",
]

RESOURCE_EXCEPTION_SOURCE_FIELDS = (
    "ExceptionName",
    "StartDateTime",
    "EndDateTime",
    "AvailabilityCode",
)


def build_resource_exception_output_objects(dataframe, source_system_owner):
    """Build ResourceException HDL records from matching Excel columns."""
    if not source_system_owner:
        raise ValueError("A source system owner is required.")

    normalized_columns = {
        normalize_column_name(column): column
        for column in dataframe.columns
    }
    missing_columns = [
        field
        for field in RESOURCE_EXCEPTION_SOURCE_FIELDS
        if normalize_column_name(field) not in normalized_columns
    ]
    if missing_columns:
        raise ValueError(
            "ResourceException requires Excel columns: "
            + ", ".join(missing_columns)
        )

    records = []
    for row_index, (_, row) in enumerate(dataframe.iterrows(), start=2):
        record = {
            field: row[normalized_columns[normalize_column_name(field)]]
            for field in RESOURCE_EXCEPTION_SOURCE_FIELDS
        }
        if all(is_empty(value) for value in record.values()):
            continue

        missing_values = [
            field for field, value in record.items() if is_empty(value)
        ]
        if missing_values:
            raise ValueError(
                f"Excel row {row_index}: ResourceException fields must not "
                "be blank: " + ", ".join(missing_values)
            )

        record["ExceptionId"] = ""
        record["SourceSystemOwner"] = source_system_owner
        record["SourceSystemId"] = f"RES_EXC_{len(records) + 1:03d}"
        records.append(record)

    return [{
        "object_name": "ResourceException",
        "metadata_fields": RESOURCE_EXCEPTION_FIELDS,
        "records": records,
    }]


def generate_resource_exception_dat(
    dataframe,
    source_system_owner,
    output_dir,
):
    """Write ResourceException.dat from Excel exception columns."""
    output_file = Path(output_dir) / "ResourceException.dat"
    return generate_hdl_bundle(
        output_file,
        build_resource_exception_output_objects(
            dataframe,
            source_system_owner,
        ),
    )
