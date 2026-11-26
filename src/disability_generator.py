import re
from pathlib import Path

from src.dat_generator import generate_hdl_bundle
from src.field_mapper import normalize_column_name
from src.validator import is_empty


PERSON_DISABILITY_FIELDS = [
    "DisabilityId",
    "DisabilityCode",
    "EffectiveStartDate",
    "EffectiveEndDate",
    "LegislationCode",
    "PersonId",
    "PersonNumber",
    "OrganizationId",
    "OrganizationName",
    "OrganizationClassificationCode",
    "RegistrationId",
    "RegistrationDate",
    "RegistrationExpDate",
    "Category",
    "Description",
    "Degree",
    "QuotaFTE",
    "Reason",
    "WorkRestriction",
    "Status",
]


NEGATIVE_DISABILITY_VALUES = {
    "no",
    "none",
    "not applicable",
    "not available",
    "n a",
    "na",
    "false",
    "0",
}

AFFIRMATIVE_DISABILITY_VALUES = {"yes", "y", "true", "1"}
DEFAULT_AFFIRMATIVE_DISABILITY_NAME = "Other disability"


def _normalized_disability_value(value):
    return re.sub(
        r"[^a-z0-9]+",
        " ",
        str(value).strip().casefold(),
    ).strip()


def _is_negative_disability_value(value):
    if is_empty(value):
        return True

    normalized = _normalized_disability_value(value)
    if not normalized:
        return True

    return (
        normalized in NEGATIVE_DISABILITY_VALUES
        or normalized.startswith("no ")
    )


def _is_affirmative_disability_value(value):
    return (
        not is_empty(value)
        and _normalized_disability_value(value)
        in AFFIRMATIVE_DISABILITY_VALUES
    )


def has_disability_data(canonical_rows):
    return any(
        not _is_negative_disability_value(row.get("Disability"))
        or not _is_negative_disability_value(row.get("DisabilityCode"))
        for row in canonical_rows
    )


def _text(value):
    if is_empty(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _first_nonempty(*values):
    return next((value for value in values if not is_empty(value)), "")


def _disability_code_lookup(lookup_rows):
    lookup = {}
    for row in lookup_rows:
        code = _text(row.get("disability_code"))
        name = _text(row.get("disability_name"))
        if not code:
            continue
        for value in (code, name):
            key = normalize_column_name(value)
            if not key:
                continue
            existing = lookup.get(key)
            if existing and existing != code:
                raise ValueError(
                    f"Disability lookup value '{value}' maps to multiple "
                    "disability codes."
                )
            lookup[key] = code
    return lookup


def _enabled_disability_code_by_name(lookup_rows, name):
    normalized_name = normalize_column_name(name)
    return next(
        (
            _text(row.get("disability_code"))
            for row in lookup_rows
            if normalize_column_name(_text(row.get("disability_name")))
            == normalized_name
            and _text(row.get("disability_code"))
        ),
        "",
    )


def _resolve_disability_code(value, code_lookup, lookup_rows):
    normalized_value = normalize_column_name(_text(value))
    if not normalized_value:
        return ""

    exact_code = code_lookup.get(normalized_value)
    if exact_code:
        return exact_code

    matches = []
    for lookup_row in lookup_rows:
        name = _text(lookup_row.get("disability_name"))
        code = _text(lookup_row.get("disability_code"))
        normalized_name = normalize_column_name(name)
        if normalized_name and normalized_name in normalized_value and code:
            matches.append((len(normalized_name), code, name))

    if not matches:
        return ""

    longest_match_length = max(match[0] for match in matches)
    longest_matches = {
        (code, name)
        for length, code, name in matches
        if length == longest_match_length
    }
    if len(longest_matches) > 1:
        names = ", ".join(sorted(name for _, name in longest_matches))
        raise ValueError(
            f"Disability details '{_text(value)}' match multiple lookup "
            f"values: {names}."
        )

    return next(iter(longest_matches))[0]


def build_person_disability_records(canonical_rows, lookup_rows):
    """Resolve populated disability values and build HDL PersonDisability rows."""
    code_lookup = _disability_code_lookup(lookup_rows)
    default_code = _enabled_disability_code_by_name(
        lookup_rows,
        DEFAULT_AFFIRMATIVE_DISABILITY_NAME,
    )
    records = []

    for row_number, row in enumerate(canonical_rows, start=2):
        disability_value = row.get("Disability")
        if _is_affirmative_disability_value(disability_value):
            disability_candidates = (
                row.get("DisabilityDetails"),
                row.get("Description"),
                row.get("DisabilityCode"),
            )
            resolved_code = next(
                (
                    code
                    for candidate in disability_candidates
                    if not _is_negative_disability_value(candidate)
                    for code in (
                        _resolve_disability_code(
                            candidate,
                            code_lookup,
                            lookup_rows,
                        ),
                    )
                    if code
                ),
                "",
            )
            has_details = any(
                not _is_negative_disability_value(value)
                for value in disability_candidates[:2]
            )
            if resolved_code:
                disability_value = resolved_code
            elif has_details:
                details = _first_nonempty(*disability_candidates[:2])
                raise ValueError(
                    f"Excel row {row_number}: disability details "
                    f"'{_text(details)}' did not match an enabled "
                    "person_disability_lookups value."
                )
            else:
                if not default_code:
                    raise ValueError(
                        f"Excel row {row_number}: Disability is marked Yes, "
                        f"but the enabled lookup does not contain "
                        f"'{DEFAULT_AFFIRMATIVE_DISABILITY_NAME}'."
                    )
                disability_value = default_code
        elif _is_negative_disability_value(disability_value):
            disability_value = row.get("DisabilityCode")
        if _is_negative_disability_value(disability_value):
            continue

        key = normalize_column_name(_text(disability_value))
        disability_code = code_lookup.get(key)
        if not disability_code:
            raise ValueError(
                f"Excel row {row_number}: disability "
                f"'{_text(disability_value)}' was not found in the enabled "
                "person_disability_lookups database values."
            )

        person_number = _text(row.get("PersonNumber"))
        if not person_number:
            raise ValueError(
                f"Excel row {row_number}: Person Number is required when "
                "disability data is provided."
            )

        effective_start_date = row.get("EffectiveStartDate")
        if is_empty(effective_start_date):
            effective_start_date = row.get("StartDate")

        records.append({
            "DisabilityId": _text(row.get("DisabilityId")),
            "DisabilityCode": disability_code,
            "EffectiveStartDate": effective_start_date,
            "EffectiveEndDate": row.get("EffectiveEndDate", ""),
            "LegislationCode": _text(row.get("LegislationCode")),
            "PersonId": _text(row.get("PersonId")),
            "PersonNumber": person_number,
            "OrganizationId": _text(row.get("OrganizationId")),
            "OrganizationName": _text(row.get("OrganizationName")),
            "OrganizationClassificationCode": _text(
                row.get("OrganizationClassificationCode")
            ),
            "RegistrationId": _text(row.get("RegistrationId")),
            "RegistrationDate": row.get("RegistrationDate", ""),
            "RegistrationExpDate": row.get("RegistrationExpDate", ""),
            "Category": _text(row.get("Category")),
            "Description": _text(_first_nonempty(
                row.get("DisabilityDetails"),
                row.get("Description"),
            )),
            "Degree": _text(row.get("Degree")),
            "QuotaFTE": _text(row.get("QuotaFTE")),
            "Reason": _text(row.get("Reason")),
            "WorkRestriction": _text(row.get("WorkRestriction")),
            "Status": _text(row.get("Status")),
        })

    return records


def generate_person_disability(
    canonical_rows,
    lookup_rows,
    output_dir,
    legislation_code="",
):
    """Write PersonDisability.dat only when the workbook has disability data."""
    output_file = Path(output_dir) / "PersonDisability.dat"
    if not has_disability_data(canonical_rows):
        output_file.unlink(missing_ok=True)
        return None

    records = build_person_disability_records(canonical_rows, lookup_rows)
    if not is_empty(legislation_code):
        for record in records:
            record["LegislationCode"] = _text(legislation_code)

    output_file.parent.mkdir(parents=True, exist_ok=True)
    return generate_hdl_bundle(
        output_file,
        [{
            "object_name": "PersonDisability",
            "metadata_fields": PERSON_DISABILITY_FIELDS,
            "records": records,
        }],
    )
