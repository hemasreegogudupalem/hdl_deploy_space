import re
from pathlib import Path

from src.dat_generator import generate_hdl_bundle
from src.field_mapper import normalize_column_name
from src.validator import format_hdl_date, is_empty


GRADE_FIELDS = [
    "SourceSystemOwner",
    "SourceSystemId",
    "EffectiveStartDate",
    "EffectiveEndDate",
    "SetCode",
    "GradeCode",
    "GradeName",
    "ActiveStatus",
]

_GRADE_SOURCE_ALIASES = {
    "EffectiveStartDate": (
        "GradeEffectiveStartDate",
        "EffectiveStartDate",
    ),
    "EffectiveEndDate": (
        "GradeEffectiveEndDate",
        "EffectiveEndDate",
    ),
    "SetCode": ("SetCode",),
    "GradeCode": ("GradeCode",),
    "GradeName": ("GradeName",),
    "ActiveStatus": ("ActiveStatus",),
}

_GRADE_REQUIRED_FIELDS = (
    "EffectiveStartDate",
    "GradeCode",
    "GradeName",
)


def build_grade_output_objects(dataframe, source_system_owner):
    """Build distinct Grade records from workbook grade columns."""
    source_system_owner = (
        str(source_system_owner).strip()
        if not is_empty(source_system_owner)
        else ""
    )

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
        for field, aliases in _GRADE_SOURCE_ALIASES.items()
    }
    missing_columns = [
        field for field in _GRADE_REQUIRED_FIELDS
        if source_columns[field] is None
    ]
    if missing_columns:
        raise ValueError(
            "Grade requires Excel columns: " + ", ".join(missing_columns)
        )

    records = []
    records_by_code = {}
    seen_source_ids = set()
    for row_index, (_, row) in enumerate(dataframe.iterrows(), start=2):
        source_values = {
            field: row[column] if column is not None else ""
            for field, column in source_columns.items()
        }
        if all(is_empty(value) for value in source_values.values()):
            continue

        missing_values = [
            field for field in _GRADE_REQUIRED_FIELDS
            if is_empty(source_values[field])
        ]
        if missing_values:
            raise ValueError(
                f"Excel row {row_index}: Grade fields must not be blank: "
                + ", ".join(missing_values)
            )

        grade_code = str(source_values["GradeCode"]).strip()
        grade_name = str(source_values["GradeName"]).strip()
        start_date = format_hdl_date(source_values["EffectiveStartDate"])
        end_date = (
            format_hdl_date(source_values["EffectiveEndDate"])
            if not is_empty(source_values["EffectiveEndDate"])
            else "4712/12/31"
        )
        set_code = (
            str(source_values["SetCode"]).strip()
            if not is_empty(source_values["SetCode"])
            else "COMMON"
        )
        active_status = (
            str(source_values["ActiveStatus"]).strip()
            if not is_empty(source_values["ActiveStatus"])
            else "A"
        )

        code_key = grade_code.casefold()
        signature = (
            start_date,
            end_date,
            set_code,
            grade_name,
            active_status,
        )
        existing = records_by_code.get(code_key)
        if existing is not None:
            if existing["signature"] != signature:
                raise ValueError(
                    f"Excel row {row_index}: conflicting Grade data for "
                    f"GradeCode '{grade_code}'."
                )
            continue

        safe_code = re.sub(
            r"[^A-Z0-9]+",
            "_",
            grade_code.upper(),
        ).strip("_")
        if not safe_code:
            raise ValueError(
                f"Excel row {row_index}: GradeCode cannot be used to "
                "generate a SourceSystemId."
            )
        source_system_id = f"GRD_{safe_code}"
        if source_system_id in seen_source_ids:
            raise ValueError(
                f"Excel row {row_index}: duplicate SourceSystemId "
                f"'{source_system_id}'."
            )
        seen_source_ids.add(source_system_id)

        record = {
            "SourceSystemOwner": source_system_owner,
            "SourceSystemId": source_system_id,
            "EffectiveStartDate": start_date,
            "EffectiveEndDate": end_date,
            "SetCode": set_code,
            "GradeCode": grade_code,
            "GradeName": grade_name,
            "ActiveStatus": active_status,
        }
        records_by_code[code_key] = {
            "signature": signature,
            "record": record,
        }
        records.append(record)

    if not records:
        raise ValueError("The Excel workbook contains no Grade records.")

    return [{
        "object_name": "Grade",
        "metadata_fields": GRADE_FIELDS,
        "records": records,
    }]


def generate_grade_dat(dataframe, source_system_owner, output_dir):
    """Write Grade.dat from distinct grade definitions in Excel."""
    return generate_hdl_bundle(
        Path(output_dir) / "Grade.dat",
        build_grade_output_objects(dataframe, source_system_owner),
    )
