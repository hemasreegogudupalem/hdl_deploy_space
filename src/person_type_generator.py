from pathlib import Path

from src.dat_generator import generate_hdl_bundle
from src.field_mapper import normalize_column_name
from src.validator import is_empty


PERSON_TYPE_FIELDS = [
    "PersonTypeId",
    "SystemPersonType",
    "UserPersonType",
    "ActiveFlag",
    "DefaultFlag",
    "SourceSystemOwner",
    "SourceSystemId",
]

_PERSON_TYPE_SOURCE_FIELDS = (
    "SystemPersonType",
    "UserPersonType",
    "ActiveFlag",
    "DefaultFlag",
)


def build_person_type_output_objects(dataframe, source_system_owner):
    """Build one PersonType record for each distinct Excel definition."""
    if not source_system_owner:
        raise ValueError("A source system owner is required.")

    normalized_columns = {
        normalize_column_name(column): column
        for column in dataframe.columns
    }
    missing_columns = [
        field for field in _PERSON_TYPE_SOURCE_FIELDS
        if normalize_column_name(field) not in normalized_columns
    ]
    if missing_columns:
        raise ValueError(
            "PersonType requires Excel columns: "
            + ", ".join(missing_columns)
        )

    records = []
    seen_definitions = set()
    for row_index, (_, row) in enumerate(dataframe.iterrows(), start=2):
        source_values = {
            field: row[normalized_columns[normalize_column_name(field)]]
            for field in _PERSON_TYPE_SOURCE_FIELDS
        }
        if all(is_empty(value) for value in source_values.values()):
            continue

        missing_values = [
            field for field, value in source_values.items() if is_empty(value)
        ]
        if missing_values:
            raise ValueError(
                f"Excel row {row_index}: PersonType fields must not be blank: "
                + ", ".join(missing_values)
            )

        system_person_type = str(source_values["SystemPersonType"]).strip()
        user_person_type = str(source_values["UserPersonType"]).strip()
        active_flag = str(source_values["ActiveFlag"]).strip()
        default_flag = str(source_values["DefaultFlag"]).strip()
        definition = (
            system_person_type,
            user_person_type,
            active_flag,
            default_flag,
        )
        if definition in seen_definitions:
            continue

        seen_definitions.add(definition)
        records.append({
            "PersonTypeId": "",
            "SystemPersonType": system_person_type,
            "UserPersonType": user_person_type,
            "ActiveFlag": active_flag,
            "DefaultFlag": default_flag,
            "SourceSystemOwner": source_system_owner,
            "SourceSystemId": "",
        })

    definitions_by_system_type = {}
    for record in records:
        system_person_type = record["SystemPersonType"]
        definitions_by_system_type.setdefault(
            system_person_type,
            [],
        ).append(record)

    for system_person_type, definitions in definitions_by_system_type.items():
        if system_person_type == "EMP":
            definitions.sort(
                key=lambda record: (
                    record["UserPersonType"] != "Permanent Employee"
                )
            )
        for sequence, record in enumerate(definitions, start=1):
            record["SourceSystemId"] = (
                f"NEW_{system_person_type}_{sequence:03d}"
            )

    return [{
        "object_name": "PersonType",
        "metadata_fields": PERSON_TYPE_FIELDS,
        "records": records,
    }]


def generate_person_type_dat(dataframe, source_system_owner, output_dir):
    """Write PersonType.dat from distinct PersonType definitions in Excel."""
    output_file = Path(output_dir) / "PersonType.dat"
    return generate_hdl_bundle(
        output_file,
        build_person_type_output_objects(dataframe, source_system_owner),
    )
