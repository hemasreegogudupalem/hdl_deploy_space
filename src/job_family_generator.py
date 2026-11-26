from pathlib import Path

from src.dat_generator import generate_hdl_bundle
from src.field_mapper import normalize_column_name
from src.validator import format_hdl_date, is_empty


JOB_FAMILY_FIELDS = [
    "JobFamilyCode",
    "JobFamilyName",
    "EffectiveStartDate",
    "EffectiveEndDate",
    "ActiveStatus",
]

JOB_SUBFAMILY_ASSOCIATION_FIELDS = [
    "JobSubFamily",
    "JobFamilyName",
    "JobFamilyCode",
    "EffectiveStartDate",
    "EffectiveEndDate",
    "ActiveStatus",
]


def build_job_family_records(dataframe):
    """Build distinct family and subfamily records from workbook columns."""
    columns = {}
    for column in dataframe.columns:
        normalized = normalize_column_name(column)
        if normalized in columns:
            raise ValueError(
                f"Excel columns '{columns[normalized]}' and '{column}' "
                "normalize to the same field name."
            )
        columns[normalized] = column

    required_sources = {
        "JobFamilyCode": ("jobfamilycode",),
        "JobFamilyName": ("jobfamilyname",),
        "EffectiveStartDate": (
            "jobfamilyeffectivestartdate",
            "effectivestartdate",
        ),
    }
    source_columns = {}
    for field, aliases in required_sources.items():
        source_columns[field] = next(
            (columns[alias] for alias in aliases if alias in columns),
            None,
        )
        if source_columns[field] is None:
            raise ValueError(
                f"Excel workbook is missing a required JobFamily column "
                f"for '{field}'."
            )

    subfamily_column = columns.get("jobsubfamily")
    family_records = {}
    association_records = {}

    for _, row in dataframe.iterrows():
        code = row[source_columns["JobFamilyCode"]]
        name = row[source_columns["JobFamilyName"]]
        start_date = row[source_columns["EffectiveStartDate"]]
        if is_empty(code) and is_empty(name) and is_empty(start_date):
            continue
        if is_empty(code) or is_empty(name) or is_empty(start_date):
            raise ValueError(
                "Each JobFamily row must include code, name, and "
                "effective start date."
            )

        start_date = format_hdl_date(start_date)
        code = str(code).strip()
        name = str(name).strip()
        family_key = (code, name, start_date)
        family_records.setdefault(family_key, {
            "JobFamilyCode": code,
            "JobFamilyName": name,
            "EffectiveStartDate": start_date,
            "EffectiveEndDate": "4712/12/31",
            "ActiveStatus": "A",
        })

        subfamily = row[subfamily_column] if subfamily_column else ""
        if not is_empty(subfamily):
            subfamily = str(subfamily).strip()
            association_key = (subfamily, name, code, start_date)
            association_records.setdefault(association_key, {
                "JobSubFamily": subfamily,
                "JobFamilyName": name,
                "JobFamilyCode": code,
                "EffectiveStartDate": start_date,
                "EffectiveEndDate": "4712/12/31",
                "ActiveStatus": "A",
            })

    if not family_records:
        raise ValueError("The Excel workbook contains no JobFamily records.")

    return list(family_records.values()), list(association_records.values())


def generate_job_family(dataframe, output_dir):
    """Write JobFamily.dat from JobFamily columns in the workbook."""
    family_records, association_records = build_job_family_records(dataframe)
    return generate_hdl_bundle(
        Path(output_dir) / "JobFamily.dat",
        [
            {
                "object_name": "JobFamily",
                "metadata_fields": JOB_FAMILY_FIELDS,
                "records": family_records,
            },
            {
                "object_name": "JobSubFamilyAssociation",
                "metadata_fields": JOB_SUBFAMILY_ASSOCIATION_FIELDS,
                "records": association_records,
            },
        ],
    )
