import re
from pathlib import Path

from src.dat_generator import generate_hdl_bundle
from src.field_mapper import normalize_column_name
from src.validator import is_empty


LEGISLATIVE_DATA_GROUP_FIELDS = [
    "Name",
    "CostAllocationIdFlexNumber",
    "LegislationCode",
    "LegislativeDataGroupId",
    "DefaultCurrencyCode",
    "DefaultCurrency",
    "StructureInstanceName",
    "Territory",
    "SourceSystemOwner",
    "SourceSystemId",
]

LEGISLATIVE_DATA_GROUP_SOURCE_FIELDS = (
    "LegislativeDataGroupName",
    "LegislationCode",
    "DefaultCurrencyCode",
    "StructureInstanceName",
)


def build_legislative_data_group_output_objects(
    dataframe,
    source_system_owner,
):
    """Build distinct LegislativeDataGroup records from workbook data."""
    if not source_system_owner:
        raise ValueError("A source system owner is required.")

    normalized_columns = {
        normalize_column_name(column): column
        for column in dataframe.columns
    }
    source_aliases = {
        "LegislativeDataGroupName": (
            "LegislativeDataGroupName",
            "Name",
        ),
        "LegislationCode": ("LegislationCode",),
        "DefaultCurrencyCode": ("DefaultCurrencyCode",),
        "StructureInstanceName": ("StructureInstanceName",),
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
        for field, aliases in source_aliases.items()
    }
    missing_columns = [
        field for field, column in source_columns.items() if column is None
    ]
    if missing_columns:
        raise ValueError(
            "LegislativeDataGroup requires Excel columns: "
            + ", ".join(missing_columns)
        )

    source_system_id_column = normalized_columns.get(
        normalize_column_name("SourceSystemId")
    )
    records = []
    seen_ldgs = set()
    ids_by_legislation_code = {}
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
                f"Excel row {row_index}: LegislativeDataGroup fields must "
                "not be blank: " + ", ".join(missing_values)
            )

        name = str(source_values["LegislativeDataGroupName"]).strip()
        legislation_code = str(source_values["LegislationCode"]).strip().upper()
        currency_code = str(source_values["DefaultCurrencyCode"]).strip().upper()
        structure_name = str(source_values["StructureInstanceName"]).strip()
        definition = (
            name.casefold(),
            legislation_code,
            currency_code,
            structure_name.casefold(),
        )
        if definition in seen_ldgs:
            continue
        seen_ldgs.add(definition)

        sequence = ids_by_legislation_code.get(legislation_code, 0) + 1
        ids_by_legislation_code[legislation_code] = sequence
        source_system_id = ""
        if source_system_id_column is not None:
            source_system_id = row[source_system_id_column]
            if not is_empty(source_system_id):
                source_system_id = str(source_system_id).strip()
            else:
                source_system_id = ""

        if not source_system_id:
            safe_code = re.sub(
                r"[^A-Z0-9]+",
                "_",
                legislation_code,
            ).strip("_")
            if not safe_code:
                raise ValueError(
                    f"Excel row {row_index}: LegislationCode cannot be used "
                    "to generate a SourceSystemId."
                )
            source_system_id = f"LDG_{safe_code}_{sequence:02d}"

        if source_system_id in seen_source_ids:
            raise ValueError(
                f"Excel row {row_index}: duplicate SourceSystemId "
                f"'{source_system_id}'."
            )
        seen_source_ids.add(source_system_id)

        records.append({
            "Name": name,
            "CostAllocationIdFlexNumber": "",
            "LegislationCode": legislation_code,
            "LegislativeDataGroupId": "",
            "DefaultCurrencyCode": currency_code,
            "DefaultCurrency": "",
            "StructureInstanceName": structure_name,
            "Territory": "",
            "SourceSystemOwner": source_system_owner,
            "SourceSystemId": source_system_id,
        })

    return [{
        "object_name": "LegislativeDataGroup",
        "metadata_fields": LEGISLATIVE_DATA_GROUP_FIELDS,
        "records": records,
    }]


def generate_legislative_data_group_dat(
    dataframe,
    source_system_owner,
    output_dir,
):
    """Write LegislativeDataGroup.dat from workbook LDG columns."""
    output_file = Path(output_dir) / "LegislativeDataGroup.dat"
    return generate_hdl_bundle(
        output_file,
        build_legislative_data_group_output_objects(
            dataframe,
            source_system_owner,
        ),
    )
