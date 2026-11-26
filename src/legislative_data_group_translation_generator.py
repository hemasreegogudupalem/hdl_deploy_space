from pathlib import Path

from src.dat_generator import generate_hdl_bundle
from src.field_mapper import normalize_column_name
from src.legislative_data_group_generator import (
    build_legislative_data_group_output_objects,
)
from src.validator import is_empty


LEGISLATIVE_DATA_GROUP_TRANSLATION_FIELDS = [
    "Language",
    "LegislativeDataGroupId",
    "Name",
    "SourceSystemOwner",
    "SourceSystemId",
    "BaseName",
]

LEGISLATIVE_DATA_GROUP_TRANSLATION_SOURCE_FIELDS = (
    "Language",
    "LegislativeDataGroupTranslatedName",
    "LegislativeDataGroupName",
    "LegislationCode",
    "DefaultCurrencyCode",
    "StructureInstanceName",
)


def build_legislative_data_group_translation_output_objects(
    dataframe,
    source_system_owner,
):
    """Build translated LDG records linked to their source LDG identifiers."""
    if not source_system_owner:
        raise ValueError("A source system owner is required.")

    normalized_columns = {
        normalize_column_name(column): column
        for column in dataframe.columns
    }
    source_aliases = {
        "Language": ("Language",),
        "LegislativeDataGroupTranslatedName": (
            "LegislativeDataGroupTranslatedName",
        ),
        "LegislativeDataGroupName": (
            "LegislativeDataGroupName",
            "BaseName",
        ),
        "LegislationCode": ("LegislationCode",),
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
    for field in ("DefaultCurrencyCode", "StructureInstanceName"):
        if normalize_column_name(field) not in normalized_columns:
            missing_columns.append(field)
    if missing_columns:
        raise ValueError(
            "LegislativeDataGroupTranslation requires Excel columns: "
            + ", ".join(missing_columns)
        )

    ldg_records = build_legislative_data_group_output_objects(
        dataframe,
        source_system_owner,
    )[0]["records"]
    source_ids = {
        (record["Name"].casefold(), record["LegislationCode"]): (
            record["SourceSystemId"]
        )
        for record in ldg_records
    }

    records = []
    seen_translations = set()
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
                f"Excel row {row_index}: LegislativeDataGroupTranslation "
                "fields must not be blank: " + ", ".join(missing_values)
            )

        language = str(source_values["Language"]).strip()
        translated_name = str(
            source_values["LegislativeDataGroupTranslatedName"]
        ).strip()
        base_name = str(source_values["LegislativeDataGroupName"]).strip()
        legislation_code = str(
            source_values["LegislationCode"]
        ).strip().upper()
        source_system_id = source_ids.get(
            (base_name.casefold(), legislation_code)
        )
        if not source_system_id:
            raise ValueError(
                f"Excel row {row_index}: no matching LegislativeDataGroup "
                f"record for {base_name!r} ({legislation_code})."
            )

        definition = (
            language.casefold(),
            base_name.casefold(),
            translated_name.casefold(),
        )
        if definition in seen_translations:
            continue
        seen_translations.add(definition)
        records.append({
            "Language": language,
            "LegislativeDataGroupId": "",
            "Name": translated_name,
            "SourceSystemOwner": source_system_owner,
            "SourceSystemId": source_system_id,
            "BaseName": base_name,
        })

    return [{
        "object_name": "LegislativeDataGroupTranslation",
        "metadata_fields": LEGISLATIVE_DATA_GROUP_TRANSLATION_FIELDS,
        "records": records,
    }]


def generate_legislative_data_group_translation_dat(
    dataframe,
    source_system_owner,
    output_dir,
):
    """Write LegislativeDataGroupTranslation.dat from LDG workbook columns."""
    output_file = Path(output_dir) / "LegislativeDataGroupTranslation.dat"
    return generate_hdl_bundle(
        output_file,
        build_legislative_data_group_translation_output_objects(
            dataframe,
            source_system_owner,
        ),
    )
