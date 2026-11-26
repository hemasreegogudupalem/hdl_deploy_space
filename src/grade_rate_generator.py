from decimal import Decimal, InvalidOperation
from pathlib import Path

from src.dat_generator import generate_hdl_bundle
from src.field_mapper import normalize_column_name
from src.validator import format_hdl_date, is_empty


GRADE_RATE_FIELDS = [
    "RateId",
    "EffectiveStartDate",
    "EffectiveEndDate",
    "LegislativeDataGroupId",
    "RateType",
    "GradeRateName",
    "CurrencyCode",
    "RateFrequency",
    "AnnualizationFactor",
    "ActiveStatus",
    "LegislativeDataGroup",
    "ActionReasonCode",
    "SourceSystemOwner",
    "SourceSystemId",
    "ProgressionRateFlag",
    "GradeLadderId",
    "GradeLadderName",
]

GRADE_RATE_VALUE_FIELDS = [
    "RateValueId",
    "EffectiveStartDate",
    "EffectiveEndDate",
    "LegislativeDataGroupId",
    "RateId(SourceSystemId)",
    "MinimumAmount",
    "MaximumAmount",
    "MidValueAmount",
    "ValueAmount",
    "GradeRateName",
    "SetCode",
    "GradeCode",
    "LegislativeDataGroup",
    "SourceSystemOwner",
    "SourceSystemId",
]

GRADE_RATE_SOURCE_ALIASES = {
    "RateId": ("GradeRateId", "RateId"),
    "EffectiveStartDate": ("GradeRateStartDate", "EffectiveStartDate"),
    "EffectiveEndDate": ("GradeRateEndDate", "EffectiveEndDate"),
    "LegislativeDataGroupId": (
        "GradeRateLegislativeDataGroupId",
        "LegislativeDataGroupId",
    ),
    "RateType": ("GradeRateRateType", "RateType"),
    "GradeRateName": ("GradeRateName",),
    "CurrencyCode": ("GradeRateCurrencyCode", "CurrencyCode"),
    "RateFrequency": ("GradeRateRateFrequency", "RateFrequency"),
    "AnnualizationFactor": (
        "GradeRateAnnualizationFactor",
        "AnnualizationFactor",
    ),
    "ActiveStatus": (
        "GradeRateActiveStatus",
        "GradeRate_ActiveStatus",
    ),
    "LegislativeDataGroup": (
        "GradeRateLegislativeDataGroup",
        "LegislativeDataGroup",
    ),
    "ActionReasonCode": ("GradeRateActionReasonCode",),
    "ProgressionRateFlag": (
        "GradeRate_ProgressionRateFlag",
        "GradeRateProgressionRateFlag",
        "ProgressionRateFlag",
    ),
    "GradeLadderId": ("GradeRateGradeLadderId", "GradeLadderId"),
    "GradeLadderName": ("GradeRateGradeLadderName", "GradeLadderName"),
    "RateValueId": ("GradeRateValueId", "RateValueId"),
    "MinimumAmount": (
        "GradeRate_MinimumAmount",
        "MinimumAmount",
    ),
    "MaximumAmount": (
        "GradeRate_MaximumAmount",
        "MaximumAmount",
    ),
    "MidValueAmount": (
        "GradeRate_MidValueAmount",
        "MidValueAmount",
    ),
    "ValueAmount": ("GradeRate_ValueAmount", "ValueAmount"),
    "SetCode": ("GradeRate_SetCode", "GradeRateSetCode"),
    "GradeCode": ("GradeRate_GradeCode", "GradeRateGradeCode", "GradeCode"),
    "ValueSourceSystemId": (
        "GradeRateValueSourceSystemId",
        "RateValueSourceSystemId",
    ),
}

_REQUIRED_FIELDS = (
    "EffectiveStartDate",
    "EffectiveEndDate",
    "RateType",
    "GradeRateName",
    "CurrencyCode",
    "RateFrequency",
    "AnnualizationFactor",
    "LegislativeDataGroup",
    "GradeCode",
    "MinimumAmount",
    "MaximumAmount",
)


def _format_number(value, field, row_index):
    if is_empty(value):
        return ""
    try:
        number = Decimal(str(value).strip())
    except (InvalidOperation, ValueError):
        raise ValueError(
            f"Excel row {row_index}: {field} must be a number."
        ) from None
    formatted = format(number, "f")
    if "." in formatted:
        formatted = formatted.rstrip("0").rstrip(".")
    return formatted


def _value(source_values, field, default=""):
    value = source_values.get(field, "")
    if is_empty(value):
        return default
    return str(value).strip()


def build_grade_rate_output_objects(dataframe, source_system_owner):
    """Build linked GradeRate and GradeRateValue records from workbook data."""
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
        for field, aliases in GRADE_RATE_SOURCE_ALIASES.items()
    }
    required_source_fields = [
        field for field in _REQUIRED_FIELDS
        if source_columns[field] is None
    ]
    if required_source_fields:
        raise ValueError(
            "GradeRate requires Excel columns: "
            + ", ".join(required_source_fields)
        )

    rate_records = []
    value_records = []
    rate_ids_by_signature = {}
    rate_signatures_by_name = {}
    value_records_by_key = {}

    for row_index, (_, row) in enumerate(dataframe.iterrows(), start=2):
        source_values = {
            field: row[column] if column is not None else ""
            for field, column in source_columns.items()
        }
        if all(
            is_empty(source_values[field])
            for field in _REQUIRED_FIELDS
        ):
            continue

        missing_values = [
            field for field in _REQUIRED_FIELDS
            if is_empty(source_values[field])
        ]
        if missing_values:
            raise ValueError(
                f"Excel row {row_index}: GradeRate fields must not be "
                "blank: " + ", ".join(missing_values)
            )

        start_date = format_hdl_date(source_values["EffectiveStartDate"])
        end_date = format_hdl_date(source_values["EffectiveEndDate"])
        annualization_factor = _format_number(
            source_values["AnnualizationFactor"],
            "AnnualizationFactor",
            row_index,
        )
        progression_flag = _value(
            source_values,
            "ProgressionRateFlag",
            "N",
        )
        if progression_flag.casefold() not in {"y", "n"}:
            raise ValueError(
                f"Excel row {row_index}: ProgressionRateFlag must be Y or N."
            )
        active_status = _value(source_values, "ActiveStatus", "A")
        rate_name = _value(source_values, "GradeRateName")
        currency = _value(source_values, "CurrencyCode")
        rate_type = _value(source_values, "RateType")
        rate_frequency = _value(source_values, "RateFrequency")
        ldg_name = _value(source_values, "LegislativeDataGroup")
        rate_signature = (
            start_date,
            end_date,
            _value(source_values, "LegislativeDataGroupId"),
            rate_type,
            rate_name,
            currency,
            rate_frequency,
            annualization_factor,
            active_status,
            ldg_name,
            _value(source_values, "ActionReasonCode"),
            progression_flag,
            _value(source_values, "GradeLadderId"),
            _value(source_values, "GradeLadderName"),
        )

        rate_name_key = rate_name.casefold()
        existing_signature = rate_signatures_by_name.get(rate_name_key)
        if (
            existing_signature is not None
            and existing_signature != rate_signature
        ):
            raise ValueError(
                f"Excel row {row_index}: conflicting GradeRate data for "
                f"GradeRateName '{rate_name}'."
            )
        rate_signatures_by_name[rate_name_key] = rate_signature

        rate_source_id = rate_ids_by_signature.get(rate_signature)
        if rate_source_id is None:
            rate_source_id = f"RATE_{len(rate_records) + 1:03d}"
            rate_ids_by_signature[rate_signature] = rate_source_id
            rate_records.append({
                "RateId": _value(source_values, "RateId"),
                "EffectiveStartDate": start_date,
                "EffectiveEndDate": end_date,
                "LegislativeDataGroupId": _value(
                    source_values,
                    "LegislativeDataGroupId",
                ),
                "RateType": rate_type,
                "GradeRateName": rate_name,
                "CurrencyCode": currency,
                "RateFrequency": rate_frequency,
                "AnnualizationFactor": annualization_factor,
                "ActiveStatus": active_status,
                "LegislativeDataGroup": ldg_name,
                "ActionReasonCode": _value(
                    source_values,
                    "ActionReasonCode",
                ),
                "SourceSystemOwner": source_system_owner,
                "SourceSystemId": rate_source_id,
                "ProgressionRateFlag": progression_flag,
                "GradeLadderId": _value(source_values, "GradeLadderId"),
                "GradeLadderName": _value(
                    source_values,
                    "GradeLadderName",
                ),
            })

        grade_code = _value(source_values, "GradeCode")
        value_key = (rate_source_id, grade_code.casefold())
        value_record = {
            "RateValueId": _value(source_values, "RateValueId"),
            "EffectiveStartDate": start_date,
            "EffectiveEndDate": end_date,
            "LegislativeDataGroupId": _value(
                source_values,
                "LegislativeDataGroupId",
            ),
            "RateId(SourceSystemId)": rate_source_id,
            "MinimumAmount": _format_number(
                source_values["MinimumAmount"],
                "MinimumAmount",
                row_index,
            ),
            "MaximumAmount": _format_number(
                source_values["MaximumAmount"],
                "MaximumAmount",
                row_index,
            ),
            "MidValueAmount": _format_number(
                source_values.get("MidValueAmount", ""),
                "MidValueAmount",
                row_index,
            ),
            "ValueAmount": _format_number(
                source_values.get("ValueAmount", ""),
                "ValueAmount",
                row_index,
            ),
            "GradeRateName": rate_name,
            "SetCode": _value(source_values, "SetCode", "COMMON"),
            "GradeCode": grade_code,
            "LegislativeDataGroup": ldg_name,
            "SourceSystemOwner": source_system_owner,
            "SourceSystemId": _value(
                source_values,
                "ValueSourceSystemId",
            ) or f"RATE_VAL_{len(value_records) + 1:03d}",
        }
        existing_value = value_records_by_key.get(value_key)
        if existing_value is not None:
            if existing_value != value_record:
                raise ValueError(
                    f"Excel row {row_index}: conflicting GradeRateValue "
                    f"data for GradeRateName '{rate_name}' and GradeCode "
                    f"'{grade_code}'."
                )
            continue
        value_records_by_key[value_key] = value_record
        value_records.append(value_record)

    if not rate_records:
        raise ValueError("The Excel workbook contains no GradeRate records.")

    return [
        {
            "object_name": "GradeRate",
            "metadata_fields": GRADE_RATE_FIELDS,
            "records": rate_records,
        },
        {
            "object_name": "GradeRateValue",
            "metadata_fields": GRADE_RATE_VALUE_FIELDS,
            "records": value_records,
        },
    ]


def generate_grade_rate_dat(dataframe, source_system_owner, output_dir):
    """Write GradeRate.dat with linked GradeRateValue records."""
    return generate_hdl_bundle(
        Path(output_dir) / "GradeRate.dat",
        build_grade_rate_output_objects(dataframe, source_system_owner),
    )
