from pathlib import Path

from src.dat_generator import generate_hdl_bundle
from src.field_mapper import normalize_column_name
from src.validator import is_empty


ACTION_REASONS_TRANSLATION_FIELDS = [
    "ActionReasonId",
    "ActionReasonCode",
    "ActionReason",
    "Language",
    "SourceSystemId",
    "SourceSystemOwner",
]

_SOURCE_ALIASES = {
    "ActionReasonId": ("ActionReasonsTranslation_ActionReasonId", "ActionReasonId"),
    "ActionReasonCode": (
        "ActionReasonsTranslation_ActionReasonCode",
        "ActionReasonCode",
    ),
    "ActionReason": (
        "ActionReasonsTranslation_ActionReason",
        "ActionReason",
    ),
    "Language": ("ActionReasonsTranslation_Language", "Language"),
    "SourceSystemId": (
        "ActionReasonsTranslation_SourceSystemId",
        "SourceSystemId",
    ),
    "SourceSystemOwner": (
        "ActionReasonsTranslation_SourceSystemOwner",
        "SourceSystemOwner",
    ),
}

_REQUIRED_FIELDS = (
    "ActionReasonCode",
    "ActionReason",
    "Language",
    "SourceSystemId",
    "SourceSystemOwner",
)


def build_action_reasons_translation_output_objects(dataframe):
    """Build ActionReasonsTranslation rows from the selected workbook."""
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
        for field, aliases in _SOURCE_ALIASES.items()
    }
    missing_columns = [
        field for field in _REQUIRED_FIELDS
        if source_columns[field] is None
    ]
    if missing_columns:
        raise ValueError(
            "ActionReasonsTranslation requires Excel columns: "
            + ", ".join(missing_columns)
        )

    records = []
    records_by_key = {}
    for row_index, (_, row) in enumerate(dataframe.iterrows(), start=2):
        source_values = {
            field: row[column] if column is not None else ""
            for field, column in source_columns.items()
        }
        if all(is_empty(value) for value in source_values.values()):
            continue

        missing_values = [
            field for field in _REQUIRED_FIELDS
            if is_empty(source_values[field])
        ]
        if missing_values:
            raise ValueError(
                f"Excel row {row_index}: ActionReasonsTranslation fields "
                "must not be blank: " + ", ".join(missing_values)
            )

        record = {
            field: (
                ""
                if is_empty(value)
                else str(value).strip()
            )
            for field, value in source_values.items()
        }
        key = (
            record["SourceSystemId"].casefold(),
            record["Language"].casefold(),
        )
        existing = records_by_key.get(key)
        if existing is not None:
            if existing != record:
                raise ValueError(
                    f"Excel row {row_index}: conflicting "
                    "ActionReasonsTranslation data for SourceSystemId "
                    f"'{record['SourceSystemId']}' and language "
                    f"'{record['Language']}'."
                )
            continue

        records_by_key[key] = record
        records.append(record)

    if not records:
        raise ValueError(
            "The Excel workbook contains no ActionReasonsTranslation rows."
        )

    return [{
        "object_name": "ActionReasonsTranslation",
        "metadata_fields": ACTION_REASONS_TRANSLATION_FIELDS,
        "records": records,
    }]


def generate_action_reasons_translation_dat(dataframe, output_dir):
    """Write ActionReasonsTranslation.dat from workbook translation rows."""
    return generate_hdl_bundle(
        Path(output_dir) / "ActionReasonsTranslation.dat",
        build_action_reasons_translation_output_objects(dataframe),
    )
