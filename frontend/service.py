"""Frontend Service Layer for Oracle Fusion HCM HDL Automation.

Integrates with existing backend modules (main.py, src/*) without modifying
any backend files.
"""

from datetime import datetime
import json
import os
from pathlib import Path
import re
import shutil
import tempfile
from typing import Any, Dict, List, Optional
from urllib.parse import quote
import zipfile
import pandas as pd
from mysql.connector import Error

# Import unmodified backend modules
from src.dat_generator import generate_hdl_bundle
from src.action_reason_generator import (
    ACTION_REASON_FIELDS,
    generate_action_reasons_dat,
)
from src.action_reasons_translation_generator import (
    ACTION_REASONS_TRANSLATION_FIELDS,
    generate_action_reasons_translation_dat,
)
from src.disability_generator import (
    PERSON_DISABILITY_FIELDS,
    generate_person_disability,
    has_disability_data,
)
from src.calendar_event_generator import (
    CALENDAR_EVENT_COVERAGE_FIELDS,
    CALENDAR_EVENT_FIELDS,
    generate_calendar_event_dat,
)
from src.contact_generator import generate_contact_dat
from src.excel_reader import read_excel
from src.actions_generator import ACTIONS_FIELDS, generate_actions_dat
from src.actions_translation_generator import (
    ACTIONS_TRANSLATION_FIELDS,
    generate_actions_translation_dat,
)
from src.resource_exception_generator import (
    RESOURCE_EXCEPTION_FIELDS,
    generate_resource_exception_dat,
)
from src.field_mapper import (
    build_alias_index,
    load_field_mappings,
    load_national_identifier_config,
    map_excel_dataframe,
    normalize_column_name,
)
from src.job_family_generator import (
    JOB_FAMILY_FIELDS,
    JOB_SUBFAMILY_ASSOCIATION_FIELDS,
    generate_job_family,
)
from src.grade_generator import GRADE_FIELDS, generate_grade_dat
from src.grade_rate_generator import (
    GRADE_RATE_FIELDS,
    GRADE_RATE_SOURCE_ALIASES,
    GRADE_RATE_VALUE_FIELDS,
    generate_grade_rate_dat,
)
from src.legislative_data_group_generator import (
    LEGISLATIVE_DATA_GROUP_FIELDS,
    generate_legislative_data_group_dat,
)
from src.legislative_data_group_translation_generator import (
    LEGISLATIVE_DATA_GROUP_TRANSLATION_FIELDS,
    LEGISLATIVE_DATA_GROUP_TRANSLATION_SOURCE_FIELDS,
    generate_legislative_data_group_translation_dat,
)
from src.person_type_generator import PERSON_TYPE_FIELDS, generate_person_type_dat
from src.mysql_storage import (
    get_active_clients,
    get_attributes,
    get_connection,
    get_objects,
    get_person_disability_lookups,
    get_templates,
    load_settings,
    load_template_configuration,
)
from src.validator import (
    is_empty,
    normalize_national_identifier,
    validate_national_identifier,
)

# Import helper functions and constants from main.py
import main as backend_main

PERSON_DISABILITY_TEMPLATE = {
    "template_id": "generated-person-disability",
    "template_name": "PersonDisability",
    "output_file_name": "PersonDisability.dat",
    "description": (
        "Person disability records generated from workbook details and "
        "enabled MySQL disability lookups."
    ),
    "active": True,
    "objects_count": 1,
    "objects": ["PersonDisability"],
}

JOB_FAMILY_TEMPLATE = {
    "template_id": "generated-job-family",
    "template_name": "JobFamily",
    "output_file_name": "JobFamily.dat",
    "description": (
        "Job family and subfamily association records generated from "
        "the selected workbook."
    ),
    "active": True,
    "objects_count": 2,
    "objects": ["JobFamily", "JobSubFamilyAssociation"],
}

GRADE_TEMPLATE = {
    "template_id": "generated-grade",
    "template_name": "Grade",
    "output_file_name": "Grade.dat",
    "description": (
        "Distinct grade definitions generated from the selected workbook."
    ),
    "active": True,
    "objects_count": 1,
    "objects": ["Grade"],
}

RESOURCE_EXCEPTION_TEMPLATE = {
    "template_id": "generated-resource-exception",
    "template_name": "ResourceException",
    "output_file_name": "ResourceException.dat",
    "description": (
        "Resource availability exception records generated from workbook "
        "exception fields."
    ),
    "active": True,
    "objects_count": 1,
    "objects": ["ResourceException"],
}

PERSON_TYPE_TEMPLATE = {
    "template_id": "generated-person-type",
    "template_name": "PersonType",
    "output_file_name": "PersonType.dat",
    "description": (
        "Person type definitions generated from distinct PersonType values "
        "in the selected workbook."
    ),
    "active": True,
    "objects_count": 1,
    "objects": ["PersonType"],
}

LEGISLATIVE_DATA_GROUP_TEMPLATE = {
    "template_id": "generated-legislative-data-group",
    "template_name": "LegislativeDataGroup",
    "output_file_name": "LegislativeDataGroup.dat",
    "description": (
        "Legislative data groups generated from distinct LDG entries in the "
        "selected workbook."
    ),
    "active": True,
    "objects_count": 1,
    "objects": ["LegislativeDataGroup"],
}

LEGISLATIVE_DATA_GROUP_TRANSLATION_TEMPLATE = {
    "template_id": "generated-legislative-data-group-translation",
    "template_name": "LegislativeDataGroupTranslation",
    "output_file_name": "LegislativeDataGroupTranslation.dat",
    "description": (
        "Translated legislative data group names generated from the selected "
        "workbook."
    ),
    "active": True,
    "objects_count": 1,
    "objects": ["LegislativeDataGroupTranslation"],
}

ACTIONS_TEMPLATE = {
    "template_id": "generated-actions",
    "template_name": "Actions",
    "output_file_name": "Actions.dat",
    "description": (
        "Action definitions generated from the selected workbook."
    ),
    "active": True,
    "objects_count": 1,
    "objects": ["Actions"],
}

CALENDAR_EVENT_TEMPLATE = {
    "template_id": "generated-calendar-event",
    "template_name": "CalendarEvent",
    "output_file_name": "CalendarEvent.dat",
    "description": (
        "Calendar events and their coverage records generated from workbook "
        "sheets."
    ),
    "active": True,
    "objects_count": 2,
    "objects": ["CalendarEvent", "CalendarEventCoverage"],
}

ACTION_REASONS_TEMPLATE = {
    "template_id": "generated-action-reasons",
    "template_name": "ActionReasons",
    "output_file_name": "ActionReasons.dat",
    "description": (
        "Action reason definitions generated from the selected workbook."
    ),
    "active": True,
    "objects_count": 1,
    "objects": ["ActionReasons"],
}

GRADE_RATE_TEMPLATE = {
    "template_id": "generated-grade-rate",
    "template_name": "GradeRate",
    "output_file_name": "GradeRate.dat",
    "description": (
        "Grade rate definitions and grade rate values generated from "
        "workbook rate data."
    ),
    "active": True,
    "objects_count": 2,
    "objects": ["GradeRate", "GradeRateValue"],
}

ACTION_REASONS_TRANSLATION_TEMPLATE = {
    "template_id": "generated-action-reasons-translation",
    "template_name": "ActionReasonsTranslation",
    "output_file_name": "ActionReasonsTranslation.dat",
    "description": (
        "Translated action reason records generated from workbook data."
    ),
    "active": True,
    "objects_count": 1,
    "objects": ["ActionReasonsTranslation"],
}

ACTIONS_TRANSLATION_TEMPLATE = {
    "template_id": "generated-actions-translation",
    "template_name": "ActionsTranslation",
    "output_file_name": "ActionsTranslation.dat",
    "description": (
        "Translated action records generated from workbook data."
    ),
    "active": True,
    "objects_count": 1,
    "objects": ["ActionsTranslation"],
}

GRADE_RATE_DEFAULT_FIELDS = {
    ("GradeRate", "RateId"),
    ("GradeRate", "LegislativeDataGroupId"),
    ("GradeRate", "ActiveStatus"),
    ("GradeRate", "ActionReasonCode"),
    ("GradeRate", "GradeLadderId"),
    ("GradeRate", "GradeLadderName"),
    ("GradeRateValue", "RateValueId"),
    ("GradeRateValue", "LegislativeDataGroupId"),
    ("GradeRateValue", "ValueAmount"),
    ("GradeRateValue", "SetCode"),
}

GRADE_RATE_GENERATED_FIELDS = {
    ("GradeRate", "SourceSystemOwner"),
    ("GradeRate", "SourceSystemId"),
    ("GradeRateValue", "RateId(SourceSystemId)"),
    ("GradeRateValue", "SourceSystemOwner"),
    ("GradeRateValue", "SourceSystemId"),
}

_BUILTIN_TEMPLATE_LAYOUTS = {
    "grade": (
        GRADE_TEMPLATE,
        GRADE_FIELDS,
        set(),
    ),
    "resourceexception": (
        RESOURCE_EXCEPTION_TEMPLATE,
        RESOURCE_EXCEPTION_FIELDS,
        {"ExceptionId"},
    ),
    "persontype": (
        PERSON_TYPE_TEMPLATE,
        PERSON_TYPE_FIELDS,
        {"PersonTypeId"},
    ),
    "legislativedatagroup": (
        LEGISLATIVE_DATA_GROUP_TEMPLATE,
        LEGISLATIVE_DATA_GROUP_FIELDS,
        {
            "CostAllocationIdFlexNumber",
            "LegislativeDataGroupId",
            "DefaultCurrency",
            "Territory",
        },
    ),
    "legislativedatagrouptranslation": (
        LEGISLATIVE_DATA_GROUP_TRANSLATION_TEMPLATE,
        LEGISLATIVE_DATA_GROUP_TRANSLATION_FIELDS,
        {"LegislativeDataGroupId"},
    ),
    "actions": (
        ACTIONS_TEMPLATE,
        ACTIONS_FIELDS,
        {"FLEX:PER_ACT_LEG_DDF", "ActionId", "ActionTypeId", "GUID"},
    ),
    "actionreasons": (
        ACTION_REASONS_TEMPLATE,
        ACTION_REASON_FIELDS,
        {"ActionReasonId"},
    ),
    "actionreasonstranslation": (
        ACTION_REASONS_TRANSLATION_TEMPLATE,
        ACTION_REASONS_TRANSLATION_FIELDS,
        {"ActionReasonId"},
    ),
    "actionstranslation": (
        ACTIONS_TRANSLATION_TEMPLATE,
        ACTIONS_TRANSLATION_FIELDS,
        {"ActionId"},
    ),
}

_BUILTIN_SOURCE_ALIASES = {
    "grade": {
        "EffectiveStartDate": (
            "GradeEffectiveStartDate",
            "EffectiveStartDate",
        ),
    },
    "legislativedatagroup": {
        "Name": ("LegislativeDataGroupName", "Name"),
    },
    "legislativedatagrouptranslation": {
        "Name": ("LegislativeDataGroupTranslatedName",),
        "BaseName": ("LegislativeDataGroupName",),
    },
    "actions": {
        "Description": ("ActionDescription", "Description"),
        "StartDate": ("ActionStartDate", "StartDate"),
        "EndDate": ("ActionEndDate", "EndDate"),
    },
    "actionreasons": {
        "StartDate": ("ActionReasonStartDate", "StartDate"),
        "EndDate": ("ActionReasonEndDate", "EndDate"),
    },
    "actionreasonstranslation": {
        "ActionReasonId": (
            "ActionReasonsTranslation_ActionReasonId",
            "ActionReasonId",
        ),
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
    },
    "actionstranslation": {
        "ActionId": ("ActionsTranslation_ActionId", "ActionId"),
        "ActionCode": ("ActionsTranslation_ActionCode", "ActionCode"),
        "ActionName": ("ActionsTranslation_ActionName", "ActionName"),
        "Description": (
            "ActionsTranslation_Description",
            "Description",
        ),
        "Language": ("ActionsTranslation_Language", "Language"),
        "SourceSystemId": (
            "ActionsTranslation_SourceSystemId",
            "SourceSystemId",
        ),
        "SourceSystemOwner": (
            "ActionsTranslation_SourceSystemOwner",
            "SourceSystemOwner",
        ),
    },
}

_BATCH_SOURCE_FIELDS = {
    "contact": [
        ("PersonNumber", True),
        ("StartDate", True),
        ("FatherName", False),
        ("MotherName", False),
        ("Spouse", False),
        ("FatherEmail", False),
        ("MotherEmail", False),
        ("SpouseEmail", False),
        ("ContactAddress", False),
        ("TownOrCity", False),
        ("Country", False),
        ("Nationality", False),
    ],
    "persondisability": [
        ("PersonNumber", True),
        ("Disability", True),
        ("DisabilityCode", False),
        ("DisabilityDetails", False),
        ("Description", False),
        ("EffectiveStartDate", False),
        ("EffectiveEndDate", False),
        ("StartDate", False),
        ("LegislationCode", False),
        ("PersonId", False),
        ("OrganizationId", False),
        ("OrganizationName", False),
        ("OrganizationClassificationCode", False),
        ("RegistrationId", False),
        ("RegistrationDate", False),
        ("RegistrationExpDate", False),
        ("Category", False),
        ("Degree", False),
        ("QuotaFTE", False),
        ("Reason", False),
        ("WorkRestriction", False),
        ("Status", False),
    ],
    "jobfamily": [
        ("JobFamilyCode", True),
        ("JobFamilyName", True),
        ("EffectiveStartDate", True),
        ("JobSubFamily", False),
    ],
    "resourceexception": [
        ("ExceptionName", True),
        ("StartDateTime", True),
        ("EndDateTime", True),
        ("AvailabilityCode", True),
    ],
    "persontype": [
        ("SystemPersonType", True),
        ("UserPersonType", True),
        ("ActiveFlag", True),
        ("DefaultFlag", True),
    ],
    "legislativedatagroup": [
        ("LegislativeDataGroupName", True),
        ("LegislationCode", True),
        ("DefaultCurrencyCode", True),
        ("StructureInstanceName", True),
    ],
    "legislativedatagrouptranslation": [
        (field, True)
        for field in LEGISLATIVE_DATA_GROUP_TRANSLATION_SOURCE_FIELDS
    ],
}

_TEMPLATE_LOAD_ORDER = (
    "legislative data group",
    "legal employer",
    "business unit",
    "organization",
    "department",
    "location",
    "job family",
    "job",
    "position",
    "grade",
    "calendar",
    "action",
    "person type",
    "worker",
    "contact",
    "person disability",
    "resource exception",
)


def _batch_template_category(template_name: str) -> str:
    normalized = normalize_column_name(template_name)
    if normalized == "persontype" or any(
        token in normalized for token in ("lookup", "reference")
    ):
        return "Reference Data"
    if normalized in {
        "worker",
        "contact",
        "persondisability",
    } or normalized.startswith("person") or any(
        token in normalized for token in ("employee", "assignment")
    ):
        return "Employee Data"
    if any(token in normalized for token in (
        "organization", "businessunit", "department", "location",
        "job", "calendar", "action", "resource", "legislativedatagroup",
    )):
        return "Organization Data"
    return "Reference Data"


def _batch_load_order(template_name: str):
    normalized_name = normalize_column_name(template_name)
    for index, template_fragment in enumerate(_TEMPLATE_LOAD_ORDER):
        if normalize_column_name(template_fragment) in normalized_name:
            return index, normalized_name
    return len(_TEMPLATE_LOAD_ORDER), normalized_name


def _is_internal_id_field(field_name: str) -> bool:
    normalized_name = normalize_column_name(field_name)
    return (
        normalized_name.endswith("id")
        and normalized_name != "sourcesystemid"
        and not normalized_name.endswith("sourcesystemid")
    )


def get_batch_template_schema(
    template_name: str,
    settings_path: str = "config/settings.json",
) -> Dict[str, Any]:
    """Return user-mappable source fields for batch generation."""
    normalized_name = normalize_column_name(template_name)
    special_fields = _BATCH_SOURCE_FIELDS.get(normalized_name)
    if special_fields is not None:
        return {
            "template_name": template_name,
            "category": _batch_template_category(template_name),
            "source_fields": [
                {
                    "mapping_key": field_name,
                    "object_name": "",
                    "field_name": field_name,
                    "required": required,
                    "target_required": required,
                    "generated": False,
                }
                for field_name, required in special_fields
            ],
        }

    template, objects = load_template_configuration(
        settings_path,
        template_name,
    )
    source_fields = []
    for obj in objects:
        mappings_by_field = {
            mapping.get("field_name"): mapping
            for mapping in obj["mappings"]
        }
        for field in obj["fields"]:
            field_name = field["field_name"]
            mapping = mappings_by_field.get(field_name) or {}
            mapping_type = str(
                mapping.get("mapping_type") or ""
            ).strip().upper()
            internal_id = _is_internal_id_field(field_name)
            generated = (
                bool(field.get("generated_field"))
                or field_name in {"SourceSystemOwner", "SourceSystemId"}
                or mapping_type == "GENERATED"
            )
            defaulted = mapping_type in {"DEFAULT", "CONSTANT", "STATIC"}
            source_fields.append({
                "mapping_key": f"{obj['object_name']}.{field_name}",
                "object_name": obj["object_name"],
                "field_name": field_name,
                "required": backend_main._is_required(
                    field.get("required_flag")
                ) and not (
                    generated or defaulted or internal_id
                ),
                "target_required": backend_main._is_required(
                    field.get("required_flag")
                ),
                "generated": generated,
                "internal_id": internal_id,
                "defaulted": defaulted,
            })

    return {
        "template_name": template.get("template_name") or template_name,
        "category": _batch_template_category(template_name),
        "source_fields": source_fields,
    }


def _output_download_url(filename: Optional[str]) -> Optional[str]:
    if not filename:
        return None
    return f"/api/reports/download/output/{quote(filename, safe='')}"


def _person_disability_validation_error(error: ValueError) -> Optional[Dict[str, Any]]:
    match = re.match(r"Excel row (\d+):\s*(.+)", str(error))
    if not match:
        return None

    return {
        "ExcelRow": int(match.group(1)),
        "Object": "PersonDisability",
        "Field": "Disability",
        "Error": match.group(2),
    }


def test_database_connection(settings_path: str = "config/settings.json") -> Dict[str, Any]:
    """Test MySQL connectivity and return detailed status."""
    try:
        settings = load_settings(settings_path)
        db_conf = settings.get("mysql", {})
        connection = get_connection(settings_path)
        cursor = connection.cursor(dictionary=True)
        cursor.execute("SELECT DATABASE() AS db_name, VERSION() AS db_version")
        info = cursor.fetchone() or {}
        cursor.close()
        connection.close()

        templates = get_templates(settings_path)
        clients = get_active_clients(settings_path)

        return {
            "status": "connected",
            "database": info.get("db_name", db_conf.get("database")),
            "version": info.get("db_version"),
            "host": db_conf.get("host"),
            "port": db_conf.get("port", 3306),
            "user": db_conf.get("user"),
            "template_count": len(templates),
            "client_count": len(clients),
            "active_client": clients[0]["client_name"] if clients else "None",
            "source_system_owner": clients[0]["source_system_owner"] if clients else "None",
        }
    except Exception as exc:
        return {
            "status": "error",
            "error": str(exc),
            "host": "127.0.0.1",
            "database": "hdl_automate",
        }


def get_available_templates(settings_path: str = "config/settings.json") -> List[Dict[str, Any]]:
    """Retrieve MySQL templates and built-in generator templates."""
    catalog = []
    try:
        templates = get_templates(settings_path)
        for t in templates:
            t_id = t.get("template_id")
            objects = get_objects(t_id, settings_path) if t_id else []
            objects = backend_main._worker_template_objects(
                t.get("template_name") or "",
                objects,
            )
            catalog.append({
                "template_id": t_id,
                "template_name": t.get("template_name"),
                "output_file_name": (
                    t.get("output_file_name")
                    or t.get("dat_file")
                    or f"{t.get('template_name')}.dat"
                ),
                "description": t.get("description"),
                "active": bool(t.get("active")),
                "category": _batch_template_category(
                    t.get("template_name") or ""
                ),
                "objects_count": len(objects),
                "objects": [obj.get("object_name") for obj in objects],
            })
    except Exception as exc:
        print(f"Failed to load database templates: {exc}")

    for builtin in (
        PERSON_DISABILITY_TEMPLATE,
        JOB_FAMILY_TEMPLATE,
        GRADE_TEMPLATE,
        RESOURCE_EXCEPTION_TEMPLATE,
        PERSON_TYPE_TEMPLATE,
        LEGISLATIVE_DATA_GROUP_TEMPLATE,
        LEGISLATIVE_DATA_GROUP_TRANSLATION_TEMPLATE,
        ACTIONS_TEMPLATE,
        CALENDAR_EVENT_TEMPLATE,
        ACTION_REASONS_TEMPLATE,
        GRADE_RATE_TEMPLATE,
        ACTION_REASONS_TRANSLATION_TEMPLATE,
        ACTIONS_TRANSLATION_TEMPLATE,
    ):
        if not any(
            item["template_name"].casefold()
            == builtin["template_name"].casefold()
            for item in catalog
        ):
            catalog.append({
                **builtin,
                "category": _batch_template_category(
                    builtin["template_name"]
                ),
            })
    return catalog


def get_template_details(template_name: str, settings_path: str = "config/settings.json") -> Dict[str, Any]:
    """Fetch deep structure of a specific template including objects and attributes."""
    if template_name.casefold() == "calendarevent":
        object_layouts = (
            (
                "CalendarEvent",
                CALENDAR_EVENT_FIELDS,
                {"TreeCode", "TreeStructureCode", "TreeVersionName"},
            ),
            (
                "CalendarEventCoverage",
                CALENDAR_EVENT_COVERAGE_FIELDS,
                {"TreeCode", "TreeStructureCode", "TreeVersionName"},
            ),
        )
        return {
            "template_name": CALENDAR_EVENT_TEMPLATE["template_name"],
            "output_file_name": CALENDAR_EVENT_TEMPLATE["output_file_name"],
            "source_system_owner": "",
            "description": CALENDAR_EVENT_TEMPLATE["description"],
            "objects": [
                {
                    "object_name": object_name,
                    "attribute_count": len(fields),
                    "attributes": [
                        {
                            "field_name": field,
                            "required": field not in optional_fields,
                            "sequence_no": index,
                            "data_type": (
                                "DATETIME"
                                if "DateTime" in field
                                else "DATE" if "Date" in field else "VARCHAR"
                            ),
                        }
                        for index, field in enumerate(fields, start=1)
                    ],
                    "mappings_count": len(fields),
                }
                for object_name, fields, optional_fields in object_layouts
            ],
            "error": None,
        }

    if template_name.casefold() == "graderate":
        object_layouts = (
            ("GradeRate", GRADE_RATE_FIELDS, {
                field for object_name, field in GRADE_RATE_DEFAULT_FIELDS
                if object_name == "GradeRate"
            }),
            ("GradeRateValue", GRADE_RATE_VALUE_FIELDS, {
                field for object_name, field in GRADE_RATE_DEFAULT_FIELDS
                if object_name == "GradeRateValue"
            }),
        )
        return {
            "template_name": GRADE_RATE_TEMPLATE["template_name"],
            "output_file_name": GRADE_RATE_TEMPLATE["output_file_name"],
            "source_system_owner": "",
            "description": GRADE_RATE_TEMPLATE["description"],
            "objects": [
                {
                    "object_name": object_name,
                    "attribute_count": len(fields),
                    "attributes": [
                        {
                            "field_name": field,
                            "required": field not in optional_fields,
                            "sequence_no": index,
                            "data_type": (
                                "DATE" if "Date" in field else "VARCHAR"
                            ),
                        }
                        for index, field in enumerate(fields, start=1)
                    ],
                    "mappings_count": len(fields),
                }
                for object_name, fields, optional_fields in object_layouts
            ],
            "error": None,
        }

    builtin_layout = _BUILTIN_TEMPLATE_LAYOUTS.get(template_name.casefold())
    if builtin_layout:
        template, fields, optional_fields = builtin_layout
        object_name = template["objects"][0]
        return {
            "template_name": template["template_name"],
            "output_file_name": template["output_file_name"],
            "source_system_owner": "",
            "description": template["description"],
            "objects": [{
                "object_name": object_name,
                "attribute_count": len(fields),
                "attributes": [
                    {
                        "field_name": field,
                        "required": field not in optional_fields,
                        "sequence_no": index,
                        "data_type": (
                            "DATETIME"
                            if "DateTime" in field
                            else "DATE" if "Date" in field else "VARCHAR"
                        ),
                    }
                    for index, field in enumerate(fields, start=1)
                ],
                "mappings_count": len(fields),
            }],
            "error": None,
        }

    if template_name.casefold() == "jobfamily":
        return {
            "template_name": JOB_FAMILY_TEMPLATE["template_name"],
            "output_file_name": JOB_FAMILY_TEMPLATE["output_file_name"],
            "source_system_owner": "",
            "description": JOB_FAMILY_TEMPLATE["description"],
            "objects": [
                {
                    "object_name": object_name,
                    "attribute_count": len(fields),
                    "attributes": [
                        {
                            "field_name": field,
                            "required": field in {
                                "JobFamilyCode",
                                "JobFamilyName",
                                "EffectiveStartDate",
                            },
                            "sequence_no": index,
                            "data_type": (
                                "DATE" if "Date" in field else "VARCHAR"
                            ),
                        }
                        for index, field in enumerate(fields, start=1)
                    ],
                    "mappings_count": len(fields),
                }
                for object_name, fields in (
                    ("JobFamily", JOB_FAMILY_FIELDS),
                    (
                        "JobSubFamilyAssociation",
                        JOB_SUBFAMILY_ASSOCIATION_FIELDS,
                    ),
                )
            ],
            "error": None,
        }

    if template_name.casefold() == "persondisability":
        return {
            "template_name": "PersonDisability",
            "output_file_name": "PersonDisability.dat",
            "source_system_owner": "",
            "description": PERSON_DISABILITY_TEMPLATE["description"],
            "objects": [{
                "object_name": "PersonDisability",
                "attribute_count": len(PERSON_DISABILITY_FIELDS),
                "attributes": [
                    {
                        "field_name": field,
                        "required": field == "PersonNumber",
                        "sequence_no": index,
                        "data_type": "VARCHAR",
                    }
                    for index, field in enumerate(
                        PERSON_DISABILITY_FIELDS,
                        start=1,
                    )
                ],
                "mappings_count": len(PERSON_DISABILITY_FIELDS),
            }],
            "error": None,
        }

    try:
        template, objects = load_template_configuration(settings_path, template_name)
    except Exception as exc:
        return {
            "template_name": template_name,
            "output_file_name": f"{template_name}.dat",
            "source_system_owner": "None",
            "description": f"Template configuration: {exc}",
            "objects": [],
            "error": str(exc),
        }

    objects = backend_main._worker_template_objects(
        template.get("template_name") or "",
        objects,
    )
    structured_objects = []
    for obj in objects:
        fields = obj.get("fields", [])
        structured_objects.append({
            "object_name": obj.get("object_name"),
            "attribute_count": len(fields),
            "attributes": [
                {
                    "field_name": f.get("field_name"),
                    "required": bool(f.get("required_flag")),
                    "sequence_no": f.get("sequence_no", 0),
                    "data_type": f.get("data_type", "VARCHAR"),
                }
                for f in fields
            ],
            "mappings_count": len(obj.get("mappings", [])),
        })
    return {
        "template_name": template.get("template_name"),
        "output_file_name": template.get("output_file_name") or template.get("dat_file"),
        "source_system_owner": template.get("source_system_owner"),
        "description": template.get("description"),
        "objects": structured_objects,
        "error": None,
    }


def list_input_files(input_dir: str = "input") -> List[Dict[str, Any]]:
    """List all available Excel files in the input directory."""
    path = Path(input_dir)
    if not path.exists():
        path.mkdir(parents=True, exist_ok=True)

    files = []
    for f in path.glob("*.*"):
        if f.suffix.lower() in (".xlsx", ".xls"):
            stat = f.stat()
            files.append({
                "filename": f.name,
                "path": str(f.as_posix()),
                "size_kb": round(stat.st_size / 1024, 1),
                "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
            })
    files.sort(key=lambda x: x["modified"], reverse=True)
    return files


def delete_input_file(filename: str, input_dir: str = "input") -> Dict[str, Any]:
    """Delete one Excel workbook from the configured input directory."""
    filename_path = Path(filename)
    if (
        not filename
        or filename_path.name != filename
        or "/" in filename
        or "\\" in filename
    ):
        raise ValueError("Provide a workbook filename, not a path.")
    if filename_path.suffix.lower() not in (".xlsx", ".xls"):
        raise ValueError("Only Excel workbooks can be deleted from the input directory.")

    input_path = Path(input_dir).resolve()
    workbook_path = input_path / filename
    if workbook_path.is_symlink():
        raise ValueError("Symbolic links cannot be deleted from the input directory.")

    resolved_path = workbook_path.resolve(strict=True)
    if resolved_path.parent != input_path or not resolved_path.is_file():
        raise ValueError("The requested file is not an Excel workbook in the input directory.")

    workbook_path.unlink()
    return {"success": True, "filename": filename}


def inspect_excel_data(
    file_path: str,
    page: int = 1,
    page_size: int = 15,
    search_query: str = "",
) -> Dict[str, Any]:
    """Read Excel workbook and provide rich paginated tabular data."""
    df = read_excel(file_path)
    total_rows = len(df)
    columns = [str(c) for c in df.columns]

    # Convert dataframe to safe records (replace NaN with "")
    df_clean = df.fillna("")

    # Ensure search_query is string
    if not isinstance(search_query, str):
        search_query = ""

    # If search query provided, filter across string representation of rows
    if search_query.strip():
        q = search_query.strip().casefold()
        mask = df_clean.astype(str).apply(lambda row: any(q in cell.casefold() for cell in row), axis=1)
        df_filtered = df_clean[mask]
    else:
        df_filtered = df_clean

    filtered_count = len(df_filtered)
    start_idx = (page - 1) * page_size
    end_idx = start_idx + page_size
    page_df = df_filtered.iloc[start_idx:end_idx]

    rows = []
    for idx, row in page_df.iterrows():
        row_dict = {"_row_number": int(idx) + 2}
        for col in columns:
            val = row[col]
            if is_empty(val):
                val = ""
            elif isinstance(val, (datetime, pd.Timestamp)):
                val = val.strftime("%Y-%m-%d")
            row_dict[col] = str(val)
        rows.append(row_dict)

    return {
        "filename": Path(file_path).name,
        "total_rows": total_rows,
        "filtered_rows": filtered_count,
        "total_pages": max(1, (filtered_count + page_size - 1) // page_size),
        "current_page": page,
        "page_size": page_size,
        "columns": columns,
        "rows": rows,
    }


def analyze_field_mappings(
    file_path: str,
    template_name: str = "Worker",
    settings_path: str = "config/settings.json",
) -> Dict[str, Any]:
    """Analyze mappings between the selected Excel file and HDL template objects."""
    if template_name.casefold() == "calendarevent":
        event_dataframe = read_excel(file_path)
        coverage_dataframe = _read_excel_sheet(
            file_path,
            "CalendarEventCoverage",
        )
        object_layouts = (
            ("CalendarEvent", CALENDAR_EVENT_FIELDS, event_dataframe),
            (
                "CalendarEventCoverage",
                CALENDAR_EVENT_COVERAGE_FIELDS,
                coverage_dataframe,
            ),
        )
        optional_fields = {
            "TreeCode",
            "TreeStructureCode",
            "TreeVersionName",
        }
        mappings = []
        for object_name, fields, dataframe in object_layouts:
            normalized_columns = {
                normalize_column_name(column): column
                for column in dataframe.columns
            }
            for field in fields:
                source_column = normalized_columns.get(
                    normalize_column_name(field),
                )
                is_optional = field in optional_fields
                mappings.append({
                    "object_name": object_name,
                    "field_name": field,
                    "excel_column": source_column or "",
                    "mapping_type": (
                        "DIRECT" if source_column
                        else "DEFAULT" if is_optional
                        else ""
                    ),
                    "status": (
                        "MAPPED" if source_column or is_optional
                        else "UNMAPPED"
                    ),
                })

        mapped_count = sum(item["status"] == "MAPPED" for item in mappings)
        mapping_types: Dict[str, int] = {}
        for item in mappings:
            mapping_type = item["mapping_type"]
            if mapping_type:
                mapping_types[mapping_type] = (
                    mapping_types.get(mapping_type, 0) + 1
                )
        return {
            "template_name": CALENDAR_EVENT_TEMPLATE["template_name"],
            "excel_file": Path(file_path).name,
            "total_fields": len(mappings),
            "mapped_count": mapped_count,
            "unmapped_count": len(mappings) - mapped_count,
            "mapping_types": mapping_types,
            "mappings": mappings,
        }

    if template_name.casefold() == "graderate":
        dataframe = read_excel(file_path)
        normalized_columns = {
            normalize_column_name(column): column
            for column in dataframe.columns
        }
        mappings = []
        for object_name, fields in (
            ("GradeRate", GRADE_RATE_FIELDS),
            ("GradeRateValue", GRADE_RATE_VALUE_FIELDS),
        ):
            for field in fields:
                mapping_key = (object_name, field)
                aliases = GRADE_RATE_SOURCE_ALIASES.get(field, (field,))
                source_column = next(
                    (
                        normalized_columns[normalize_column_name(alias)]
                        for alias in aliases
                        if normalize_column_name(alias) in normalized_columns
                    ),
                    None,
                )
                if mapping_key in GRADE_RATE_GENERATED_FIELDS:
                    mapping_type = "GENERATED"
                    status = "MAPPED"
                elif mapping_key in GRADE_RATE_DEFAULT_FIELDS:
                    mapping_type = "DEFAULT"
                    status = "MAPPED"
                elif source_column:
                    mapping_type = "DIRECT"
                    status = "MAPPED"
                else:
                    mapping_type = ""
                    status = "UNMAPPED"
                mappings.append({
                    "object_name": object_name,
                    "field_name": field,
                    "excel_column": source_column or "",
                    "mapping_type": mapping_type,
                    "status": status,
                })

        mapping_types: Dict[str, int] = {}
        for item in mappings:
            mapping_type = item["mapping_type"]
            if mapping_type:
                mapping_types[mapping_type] = (
                    mapping_types.get(mapping_type, 0) + 1
                )
        mapped_count = sum(item["status"] == "MAPPED" for item in mappings)
        return {
            "template_name": GRADE_RATE_TEMPLATE["template_name"],
            "excel_file": Path(file_path).name,
            "total_fields": len(mappings),
            "mapped_count": mapped_count,
            "unmapped_count": len(mappings) - mapped_count,
            "mapping_types": mapping_types,
            "mappings": mappings,
        }

    builtin_layout = _BUILTIN_TEMPLATE_LAYOUTS.get(template_name.casefold())
    if builtin_layout:
        template, fields, optional_fields = builtin_layout
        dataframe = read_excel(file_path)
        normalized_columns = {
            normalize_column_name(column): column
            for column in dataframe.columns
        }
        mappings = []
        aliases_by_field = _BUILTIN_SOURCE_ALIASES.get(
            template["template_name"].casefold(),
            {},
        )
        for field in fields:
            aliases = aliases_by_field.get(field, (field,))
            source_column = next(
                (
                    normalized_columns[normalize_column_name(alias)]
                    for alias in aliases
                    if normalize_column_name(alias) in normalized_columns
                ),
                None,
            )
            if (
                field in {"SourceSystemOwner", "SourceSystemId"}
                and field not in aliases_by_field
            ):
                mapping_type = "GENERATED"
                status = "MAPPED"
            elif source_column:
                mapping_type = "DIRECT"
                status = "MAPPED"
            elif field in optional_fields:
                mapping_type = "DEFAULT"
                status = "MAPPED"
            else:
                mapping_type = ""
                status = "UNMAPPED"
            mappings.append({
                "object_name": template["objects"][0],
                "field_name": field,
                "excel_column": source_column or "",
                "mapping_type": mapping_type,
                "status": status,
            })

        mapped_count = sum(item["status"] == "MAPPED" for item in mappings)
        mapping_types: Dict[str, int] = {}
        for item in mappings:
            mapping_type = item["mapping_type"]
            if mapping_type:
                mapping_types[mapping_type] = (
                    mapping_types.get(mapping_type, 0) + 1
                )
        return {
            "template_name": template["template_name"],
            "excel_file": Path(file_path).name,
            "total_fields": len(mappings),
            "mapped_count": mapped_count,
            "unmapped_count": len(mappings) - mapped_count,
            "mapping_types": mapping_types,
            "mappings": mappings,
        }

    if template_name.casefold() == "jobfamily":
        df = read_excel(file_path)
        normalized_columns = {
            normalize_column_name(column): column
            for column in df.columns
        }
        source_fields = {
            ("JobFamily", "JobFamilyCode"): ("JobFamilyCode",),
            ("JobFamily", "JobFamilyName"): ("JobFamilyName",),
            ("JobFamily", "EffectiveStartDate"): (
                "JobFamilyEffectiveStartDate",
                "EffectiveStartDate",
            ),
            ("JobSubFamilyAssociation", "JobSubFamily"): ("JobSubFamily",),
        }
        mappings = []
        for object_name, fields in (
            ("JobFamily", JOB_FAMILY_FIELDS),
            ("JobSubFamilyAssociation", JOB_SUBFAMILY_ASSOCIATION_FIELDS),
        ):
            for field in fields:
                aliases = source_fields.get((object_name, field), ())
                source_column = next(
                    (
                        normalized_columns[normalize_column_name(alias)]
                        for alias in aliases
                        if normalize_column_name(alias) in normalized_columns
                    ),
                    "",
                )
                mappings.append({
                    "object_name": object_name,
                    "field_name": field,
                    "excel_column": source_column,
                    "mapping_type": "DIRECT" if source_column else "DEFAULT",
                    "status": "MAPPED" if source_column else "UNMAPPED",
                })
        mapped_count = sum(
            item["status"] == "MAPPED" for item in mappings
        )
        return {
            "template_name": JOB_FAMILY_TEMPLATE["template_name"],
            "excel_file": Path(file_path).name,
            "total_fields": len(mappings),
            "mapped_count": mapped_count,
            "unmapped_count": len(mappings) - mapped_count,
            "mapping_types": {"DIRECT": mapped_count},
            "mappings": mappings,
        }

    if template_name.casefold() == "persondisability":
        df = read_excel(file_path)
        alias_map = load_field_mappings()
        normalized_columns = {
            normalize_column_name(column): column
            for column in df.columns
        }
        mapping_report = []
        for field in PERSON_DISABILITY_FIELDS:
            aliases = [field, *alias_map.get(field, [])]
            source_column = next(
                (
                    normalized_columns[normalize_column_name(alias)]
                    for alias in aliases
                    if normalize_column_name(alias) in normalized_columns
                ),
                None,
            )
            mapping_report.append({
                "object_name": "PersonDisability",
                "field_name": field,
                "excel_column": source_column or "",
                "mapping_type": "COLUMN" if source_column else "",
                "status": "MAPPED" if source_column else "UNMAPPED",
            })

        mapped_count = sum(
            item["status"] == "MAPPED" for item in mapping_report
        )
        return {
            "template_name": "PersonDisability",
            "excel_file": Path(file_path).name,
            "total_fields": len(mapping_report),
            "mapped_count": mapped_count,
            "unmapped_count": len(mapping_report) - mapped_count,
            "mapping_types": {"COLUMN": mapped_count},
            "mappings": mapping_report,
        }

    template, objects = load_template_configuration(settings_path, template_name)
    objects = backend_main._worker_template_objects(
        template.get("template_name") or "",
        objects,
    )
    df = read_excel(file_path)
    alias_map = load_field_mappings()
    alias_index = build_alias_index(alias_map)

    normalized_columns = {}
    for column in df.columns:
        normalized = normalize_column_name(column)
        normalized_columns[normalized] = column

    mapping_lookup, mapping_report = backend_main._make_mapping_lookup(
        objects,
        normalized_columns,
        alias_index,
    )

    # Summaries
    mapped_count = sum(1 for m in mapping_report if m["status"] == "MAPPED")
    unmapped_count = sum(1 for m in mapping_report if m["status"] == "UNMAPPED")
    type_counts: Dict[str, int] = {}
    for m in mapping_report:
        mtype = m.get("mapping_type") or "NONE"
        type_counts[mtype] = type_counts.get(mtype, 0) + 1

    return {
        "template_name": template["template_name"],
        "excel_file": Path(file_path).name,
        "total_fields": len(mapping_report),
        "mapped_count": mapped_count,
        "unmapped_count": unmapped_count,
        "mapping_types": type_counts,
        "mappings": mapping_report,
    }


def _normalize_job_records(
    records: List[Dict[str, Any]],
    source_owner: str,
) -> List[Dict[str, Any]]:
    jobs = {}
    for record in records:
        name = str(record.get("Name", "")).strip()
        if not name:
            raise ValueError("Job Name is required to generate Job records.")

        job_code = re.sub(r"[^A-Z0-9]+", "_", name.upper()).strip("_")
        if not job_code:
            raise ValueError(f"Cannot derive a JobCode from job name {name!r}.")

        existing = jobs.get(job_code)
        if existing and existing["Name"].casefold() != name.casefold():
            raise ValueError(f"JobCode collision for job name {name!r}.")

        record = {
            **record,
            "JobCode": job_code,
            "SetCode": record.get("SetCode") or "COMMON",
            "ActiveStatus": record.get("ActiveStatus") or "A",
            "SourceSystemOwner": source_owner,
            "SourceSystemId": f"JOB_{job_code}",
        }
        effective_start_date = pd.to_datetime(
            record.get("EffectiveStartDate"),
            errors="coerce",
        )
        if pd.isna(effective_start_date):
            raise ValueError(
                f"Job {name!r} has an invalid EffectiveStartDate."
            )
        record["EffectiveStartDate"] = effective_start_date

        if (
            existing is None
            or effective_start_date
            < pd.to_datetime(existing["EffectiveStartDate"])
        ):
            jobs[job_code] = record

    return list(jobs.values())


def detect_generation_defaults(
    file_path: str,
    template_name: str = "Worker",
    settings_path: str = "config/settings.json",
) -> Dict[str, Any]:
    """Auto-detect common generation parameter values from Excel data."""
    if template_name.casefold() in _BUILTIN_TEMPLATE_LAYOUTS:
        clients = get_active_clients(settings_path)
        if len(clients) != 1:
            raise ValueError(
                "Expected exactly one active client for HDL source identifiers; "
                f"found {len(clients)}."
            )
        return {
            "legislation_code": "",
            "legal_employer_name": "",
            "source_system_owner": (
                clients[0].get("source_system_owner") or ""
            ),
        }

    if template_name.casefold() == "jobfamily":
        return {
            "legislation_code": "",
            "legal_employer_name": "",
            "source_system_owner": "",
        }

    config_template_name = (
        "Worker"
        if template_name.casefold() == "persondisability"
        else template_name
    )
    template, _ = load_template_configuration(
        settings_path,
        config_template_name,
    )
    df = read_excel(file_path)
    alias_map = load_field_mappings()
    canonical_rows, _ = map_excel_dataframe(df, alias_map)

    defaults = backend_main._worker_generation_defaults(
        canonical_rows,
        template["source_system_owner"],
    )

    return {
        "legislation_code": defaults.get("legislation_code") or "",
        "legal_employer_name": defaults.get("legal_employer_name") or "",
        "source_system_owner": defaults.get("source_system_owner") or template.get("source_system_owner") or "GRISMO",
    }


def _run_person_disability_pipeline(
    file_path: str,
    generation_overrides: Dict[str, str],
    settings_path: str,
    output_dir: str,
    column_mappings=None,
) -> Dict[str, Any]:
    worker_template, _ = load_template_configuration(
        settings_path,
        "Worker",
    )
    df = read_excel(file_path)
    alias_map = load_field_mappings()
    canonical_rows, _ = map_excel_dataframe(df, alias_map)
    canonical_rows = _apply_canonical_column_mappings(
        df,
        canonical_rows,
        column_mappings,
    )

    defaults = backend_main._worker_generation_defaults(
        canonical_rows,
        worker_template["source_system_owner"],
    )
    legislation_code = (
        generation_overrides.get("legislation_code")
        or defaults.get("legislation_code")
        or ""
    )
    legislation_code = backend_main._normalize_legislation_code(
        legislation_code
    )

    disability_lookups = (
        get_person_disability_lookups(settings_path)
        if has_disability_data(canonical_rows)
        else []
    )
    disability_output_file = generate_person_disability(
        canonical_rows,
        disability_lookups,
        output_dir,
        legislation_code=legislation_code,
    )
    if disability_output_file is None:
        raise ValueError(
            "The selected workbook has no PersonDisability records to generate."
        )

    lines = disability_output_file.read_text(
        encoding="utf-8",
    ).splitlines()
    generated_rows = sum(
        line.startswith("MERGE|PersonDisability|")
        for line in lines
    )
    output_path = disability_output_file.as_posix()
    return {
        "success": True,
        "template_name": "PersonDisability",
        "output_file": output_path,
        "output_filename": disability_output_file.name,
        "output_download_url": _output_download_url(
            disability_output_file.name
        ),
        "output_file_size_kb": round(
            disability_output_file.stat().st_size / 1024,
            2,
        ),
        "contact_output_file": None,
        "person_disability_output_file": output_path,
        "person_disability_output_filename": disability_output_file.name,
        "person_disability_download_url": _output_download_url(
            disability_output_file.name
        ),
        "total_excel_rows": len(df),
        "generated_rows": generated_rows,
        "skipped_rows": 0,
        "validation_errors_count": 0,
        "validation_errors": [],
        "mapping_report_path": "",
        "validation_report_path": "",
        "dat_preview": lines[:60],
        "generation_values": {
            "legislation_code": legislation_code,
            "legal_employer_name": "",
            "source_system_owner": (
                generation_overrides.get("source_system_owner")
                or defaults.get("source_system_owner")
                or worker_template.get("source_system_owner")
                or ""
            ),
        },
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def _mapped_source_dataframe(dataframe, column_mappings):
    mapped_dataframe = dataframe.copy()
    mapped_sources = {}
    for mapping_key, source_column in (column_mappings or {}).items():
        if not source_column:
            continue
        if source_column not in dataframe.columns:
            raise ValueError(
                f"Mapped Excel column '{source_column}' for "
                f"'{mapping_key}' does not exist in the selected worksheet."
            )
        field_name = mapping_key.rsplit(".", 1)[-1]
        existing_source = mapped_sources.get(field_name)
        if (
            existing_source is not None
            and existing_source != source_column
        ):
            raise ValueError(
                f"Multiple source columns were selected for '{field_name}'."
            )
        mapped_sources[field_name] = source_column
        mapped_dataframe[field_name] = dataframe[source_column]
    return mapped_dataframe


def _read_excel_sheet(file_path, sheet_name):
    dataframe = pd.read_excel(file_path, sheet_name=sheet_name)
    dataframe = dataframe.dropna(axis=1, how="all")
    dataframe.columns = [str(column).strip() for column in dataframe.columns]
    return dataframe


def _mapped_calendar_event_dataframes(
    event_dataframe,
    coverage_dataframe,
    column_mappings,
):
    event_mappings = {}
    coverage_mappings = {}
    for mapping_key, source_column in (column_mappings or {}).items():
        if not source_column:
            continue
        if "." not in mapping_key:
            raise ValueError(
                f"CalendarEvent mapping '{mapping_key}' must include its "
                "object name."
            )
        object_name, field_name = mapping_key.split(".", 1)
        if object_name.casefold() == "calendarevent":
            event_mappings[f"CalendarEvent.{field_name}"] = source_column
        elif object_name.casefold() == "calendareventcoverage":
            coverage_mappings[
                f"CalendarEventCoverage.{field_name}"
            ] = source_column
        else:
            raise ValueError(
                f"Unknown CalendarEvent mapping object '{object_name}'."
            )

    return (
        _mapped_source_dataframe(event_dataframe, event_mappings),
        _mapped_source_dataframe(coverage_dataframe, coverage_mappings),
    )


def _apply_canonical_column_mappings(
    dataframe,
    canonical_rows,
    column_mappings,
):
    for mapping_key, source_column in (column_mappings or {}).items():
        if not source_column:
            continue
        if source_column not in dataframe.columns:
            raise ValueError(
                f"Mapped Excel column '{source_column}' for "
                f"'{mapping_key}' does not exist in the selected worksheet."
            )
        field_name = mapping_key.rsplit(".", 1)[-1]
        for row_index, value in enumerate(dataframe[source_column]):
            canonical_rows[row_index][field_name] = value
    return canonical_rows


def _run_contact_pipeline(
    file_path: str,
    generation_overrides: Dict[str, str],
    settings_path: str,
    output_dir: str,
    reports_dir: str,
    column_mappings=None,
) -> Dict[str, Any]:
    template, _ = load_template_configuration(settings_path, "Contact")
    df = read_excel(file_path)
    source_system_owner = (
        generation_overrides.get("source_system_owner")
        or template.get("source_system_owner")
        or ""
    )
    output_file = generate_contact_dat(
        df,
        source_system_owner,
        output_dir,
        source_column_mappings=column_mappings,
    )
    lines = output_file.read_text(encoding="utf-8").splitlines()
    generated_rows = sum(line.startswith("MERGE|") for line in lines)
    if generated_rows == 0:
        raise ValueError(
            "No Contact records were generated. Verify the workbook includes "
            "Person Number, Start Date, and at least one family contact name "
            "(Father Name, Mother Name, or Spouse)."
        )

    report_dir = Path(reports_dir)
    report_dir.mkdir(parents=True, exist_ok=True)
    mapping_report_path = report_dir / "mapping_report.xlsx"
    validation_report_path = report_dir / "validation_errors.csv"
    pd.DataFrame(columns=[
        "object_name",
        "field_name",
        "excel_column",
        "mapping_type",
        "default_value",
        "status",
    ]).to_excel(mapping_report_path, index=False)
    pd.DataFrame(columns=[
        "ExcelRow",
        "Object",
        "Field",
        "Error",
    ]).to_csv(validation_report_path, index=False)

    output_path = output_file.as_posix()
    return {
        "success": True,
        "template_name": template.get("template_name") or "Contact",
        "output_file": output_path,
        "output_filename": output_file.name,
        "output_download_url": _output_download_url(output_file.name),
        "output_file_size_kb": round(output_file.stat().st_size / 1024, 2),
        "contact_output_file": output_path,
        "contact_output_filename": output_file.name,
        "contact_download_url": _output_download_url(output_file.name),
        "person_disability_output_file": None,
        "person_disability_output_filename": None,
        "person_disability_download_url": None,
        "total_excel_rows": len(df),
        "generated_rows": generated_rows,
        "skipped_rows": 0,
        "validation_errors_count": 0,
        "validation_errors": [],
        "mapping_report_path": mapping_report_path.as_posix(),
        "validation_report_path": validation_report_path.as_posix(),
        "dat_preview": lines[:60],
        "generation_values": {
            "legislation_code": "",
            "legal_employer_name": "",
            "source_system_owner": source_system_owner,
        },
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def _run_resource_exception_pipeline(
    file_path: str,
    generation_overrides: Dict[str, str],
    settings_path: str,
    output_dir: str,
    reports_dir: str,
    column_mappings=None,
) -> Dict[str, Any]:
    source_system_owner = generation_overrides.get("source_system_owner")
    if not source_system_owner:
        clients = get_active_clients(settings_path)
        if len(clients) != 1:
            raise ValueError(
                "Expected exactly one active client for HDL source identifiers; "
                f"found {len(clients)}."
            )
        source_system_owner = clients[0].get("source_system_owner")
    if not source_system_owner:
        raise ValueError(
            "The active client has no source_system_owner configured."
        )

    dataframe = read_excel(file_path)
    output_file = generate_resource_exception_dat(
        _mapped_source_dataframe(dataframe, column_mappings),
        source_system_owner,
        output_dir,
    )
    generated_rows = sum(
        line.startswith("MERGE|")
        for line in output_file.read_text(encoding="utf-8").splitlines()
    )
    if generated_rows == 0:
        raise ValueError(
            "No ResourceException records were generated. Verify the workbook "
            "contains exception data."
        )

    report_dir = Path(reports_dir)
    report_dir.mkdir(parents=True, exist_ok=True)
    mapping_report_path = report_dir / "mapping_report.xlsx"
    validation_report_path = report_dir / "validation_errors.csv"
    pd.DataFrame(columns=[
        "object_name",
        "field_name",
        "excel_column",
        "mapping_type",
        "default_value",
        "status",
    ]).to_excel(mapping_report_path, index=False)
    pd.DataFrame(columns=[
        "ExcelRow",
        "Object",
        "Field",
        "Error",
    ]).to_csv(validation_report_path, index=False)

    output_path = output_file.as_posix()
    return {
        "success": True,
        "template_name": "ResourceException",
        "output_file": output_path,
        "output_filename": output_file.name,
        "output_download_url": _output_download_url(output_file.name),
        "output_file_size_kb": round(output_file.stat().st_size / 1024, 2),
        "contact_output_file": None,
        "contact_output_filename": None,
        "contact_download_url": None,
        "person_disability_output_file": None,
        "person_disability_output_filename": None,
        "person_disability_download_url": None,
        "total_excel_rows": len(dataframe),
        "generated_rows": generated_rows,
        "skipped_rows": len(dataframe) - generated_rows,
        "validation_errors_count": 0,
        "validation_errors": [],
        "mapping_report_path": mapping_report_path.as_posix(),
        "validation_report_path": validation_report_path.as_posix(),
        "dat_preview": output_file.read_text(
            encoding="utf-8",
        ).splitlines()[:60],
        "generation_values": {
            "legislation_code": "",
            "legal_employer_name": "",
            "source_system_owner": source_system_owner,
        },
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def _run_actions_pipeline(
    file_path: str,
    generation_overrides: Dict[str, str],
    settings_path: str,
    output_dir: str,
    reports_dir: str,
    column_mappings=None,
) -> Dict[str, Any]:
    source_system_owner = generation_overrides.get("source_system_owner")
    if not source_system_owner:
        clients = get_active_clients(settings_path)
        if len(clients) != 1:
            raise ValueError(
                "Expected exactly one active client for HDL source identifiers; "
                f"found {len(clients)}."
            )
        source_system_owner = clients[0].get("source_system_owner")
    if not source_system_owner:
        raise ValueError(
            "The active client has no source_system_owner configured."
        )

    dataframe = read_excel(file_path)
    output_file = generate_actions_dat(
        _mapped_source_dataframe(dataframe, column_mappings),
        source_system_owner,
        output_dir,
    )
    lines = output_file.read_text(encoding="utf-8").splitlines()
    generated_rows = sum(line.startswith("MERGE|") for line in lines)
    if generated_rows == 0:
        raise ValueError(
            "No Actions records were generated. Verify the workbook contains "
            "action code, name, description, type code, and date columns."
        )

    report_dir = Path(reports_dir)
    report_dir.mkdir(parents=True, exist_ok=True)
    mapping_report_path = report_dir / "mapping_report.xlsx"
    validation_report_path = report_dir / "validation_errors.csv"
    pd.DataFrame(columns=[
        "object_name",
        "field_name",
        "excel_column",
        "mapping_type",
        "default_value",
        "status",
    ]).to_excel(mapping_report_path, index=False)
    pd.DataFrame(columns=[
        "ExcelRow",
        "Object",
        "Field",
        "Error",
    ]).to_csv(validation_report_path, index=False)

    return {
        "success": True,
        "template_name": ACTIONS_TEMPLATE["template_name"],
        "output_file": output_file.as_posix(),
        "output_filename": output_file.name,
        "output_download_url": _output_download_url(output_file.name),
        "output_file_size_kb": round(output_file.stat().st_size, 2),
        "contact_output_file": None,
        "contact_output_filename": None,
        "contact_download_url": None,
        "person_disability_output_file": None,
        "person_disability_output_filename": None,
        "person_disability_download_url": None,
        "total_excel_rows": len(dataframe),
        "generated_rows": generated_rows,
        "skipped_rows": len(dataframe) - generated_rows,
        "validation_errors_count": 0,
        "validation_errors": [],
        "mapping_report_path": mapping_report_path.as_posix(),
        "validation_report_path": validation_report_path.as_posix(),
        "dat_preview": lines[:60],
        "generation_values": {
            "legislation_code": "",
            "legal_employer_name": "",
            "source_system_owner": source_system_owner,
        },
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def _write_empty_generation_reports(reports_dir):
    report_dir = Path(reports_dir)
    report_dir.mkdir(parents=True, exist_ok=True)
    mapping_report_path = report_dir / "mapping_report.xlsx"
    validation_report_path = report_dir / "validation_errors.csv"
    pd.DataFrame(columns=[
        "object_name",
        "field_name",
        "excel_column",
        "mapping_type",
        "default_value",
        "status",
    ]).to_excel(mapping_report_path, index=False)
    pd.DataFrame(columns=[
        "ExcelRow",
        "Object",
        "Field",
        "Error",
    ]).to_csv(validation_report_path, index=False)
    return mapping_report_path, validation_report_path


def _builtin_generation_result(
    template,
    output_file,
    total_excel_rows,
    generated_rows,
    reports_dir,
    source_system_owner="",
):
    mapping_report_path, validation_report_path = (
        _write_empty_generation_reports(reports_dir)
    )
    lines = output_file.read_text(encoding="utf-8").splitlines()
    return {
        "success": True,
        "template_name": template["template_name"],
        "output_file": output_file.as_posix(),
        "output_filename": output_file.name,
        "output_download_url": _output_download_url(output_file.name),
        "output_file_size_kb": round(output_file.stat().st_size / 1024, 2),
        "contact_output_file": None,
        "contact_output_filename": None,
        "contact_download_url": None,
        "person_disability_output_file": None,
        "person_disability_output_filename": None,
        "person_disability_download_url": None,
        "total_excel_rows": total_excel_rows,
        "generated_rows": generated_rows,
        "skipped_rows": max(total_excel_rows - generated_rows, 0),
        "validation_errors_count": 0,
        "validation_errors": [],
        "mapping_report_path": mapping_report_path.as_posix(),
        "validation_report_path": validation_report_path.as_posix(),
        "dat_preview": lines[:60],
        "generation_values": {
            "legislation_code": "",
            "legal_employer_name": "",
            "source_system_owner": source_system_owner,
        },
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def _run_calendar_event_pipeline(
    file_path,
    output_dir,
    reports_dir,
    column_mappings=None,
):
    event_dataframe = read_excel(file_path)
    coverage_dataframe = _read_excel_sheet(
        file_path,
        "CalendarEventCoverage",
    )
    event_dataframe, coverage_dataframe = (
        _mapped_calendar_event_dataframes(
            event_dataframe,
            coverage_dataframe,
            column_mappings,
        )
    )
    output_file = generate_calendar_event_dat(
        event_dataframe,
        coverage_dataframe,
        output_dir,
    )
    generated_rows = sum(
        line.startswith("MERGE|")
        for line in output_file.read_text(encoding="utf-8").splitlines()
    )
    if generated_rows == 0:
        raise ValueError(
            "No CalendarEvent records were generated. Verify the workbook "
            "contains CalendarEvent and CalendarEventCoverage sheets."
        )
    return _builtin_generation_result(
        CALENDAR_EVENT_TEMPLATE,
        output_file,
        len(event_dataframe) + len(coverage_dataframe),
        generated_rows,
        reports_dir,
    )


def _run_action_reasons_pipeline(
    file_path,
    generation_overrides,
    settings_path,
    output_dir,
    reports_dir,
    column_mappings=None,
):
    source_system_owner = generation_overrides.get("source_system_owner")
    if not source_system_owner:
        clients = get_active_clients(settings_path)
        if len(clients) != 1:
            raise ValueError(
                "Expected exactly one active client for HDL source identifiers; "
                f"found {len(clients)}."
            )
        source_system_owner = clients[0].get("source_system_owner")
    if not source_system_owner:
        raise ValueError(
            "The active client has no source_system_owner configured."
        )

    dataframe = read_excel(file_path)
    output_file = generate_action_reasons_dat(
        _mapped_source_dataframe(dataframe, column_mappings),
        source_system_owner,
        output_dir,
    )
    generated_rows = sum(
        line.startswith("MERGE|")
        for line in output_file.read_text(encoding="utf-8").splitlines()
    )
    if generated_rows == 0:
        raise ValueError(
            "No ActionReasons records were generated. Verify the workbook "
            "contains action-reason code, name, and date columns."
        )
    return _builtin_generation_result(
        ACTION_REASONS_TEMPLATE,
        output_file,
        len(dataframe),
        generated_rows,
        reports_dir,
        source_system_owner,
    )


def _mapped_translation_source_dataframe(
    dataframe,
    column_mappings,
    object_name,
    source_prefix,
):
    mapped_dataframe = dataframe.copy()
    for mapping_key, source_column in (column_mappings or {}).items():
        if not source_column:
            continue
        mapped_object, field_name = mapping_key.split(".", 1)
        if mapped_object.casefold() != object_name.casefold():
            continue
        if source_column not in dataframe.columns:
            raise ValueError(
                f"Mapped Excel column '{source_column}' for "
                f"'{mapping_key}' does not exist in the selected worksheet."
            )
        mapped_dataframe[f"{source_prefix}{field_name}"] = (
            dataframe[source_column]
        )
    return mapped_dataframe


def _run_actions_translation_pipeline(
    file_path,
    output_dir,
    reports_dir,
    column_mappings=None,
):
    dataframe = read_excel(file_path)
    dataframe = _mapped_translation_source_dataframe(
        dataframe,
        column_mappings,
        "ActionsTranslation",
        "ActionsTranslation_",
    )
    output_file = generate_actions_translation_dat(dataframe, output_dir)
    generated_rows = sum(
        line.startswith("MERGE|")
        for line in output_file.read_text(encoding="utf-8").splitlines()
    )
    if generated_rows == 0:
        raise ValueError(
            "No ActionsTranslation rows were generated. Verify the workbook "
            "contains translated action names and descriptions."
        )
    return _builtin_generation_result(
        ACTIONS_TRANSLATION_TEMPLATE,
        output_file,
        len(dataframe),
        generated_rows,
        reports_dir,
    )


def _run_action_reasons_translation_pipeline(
    file_path,
    output_dir,
    reports_dir,
    column_mappings=None,
):
    dataframe = read_excel(file_path)
    dataframe = _mapped_translation_source_dataframe(
        dataframe,
        column_mappings,
        "ActionReasonsTranslation",
        "ActionReasonsTranslation_",
    )
    output_file = generate_action_reasons_translation_dat(
        dataframe,
        output_dir,
    )
    generated_rows = sum(
        line.startswith("MERGE|")
        for line in output_file.read_text(encoding="utf-8").splitlines()
    )
    if generated_rows == 0:
        raise ValueError(
            "No ActionReasonsTranslation rows were generated. Verify the "
            "workbook contains translated action reason data."
        )
    return _builtin_generation_result(
        ACTION_REASONS_TRANSLATION_TEMPLATE,
        output_file,
        len(dataframe),
        generated_rows,
        reports_dir,
    )


def _mapped_grade_rate_source_dataframe(dataframe, column_mappings):
    mapped_dataframe = dataframe.copy()
    mapped_sources = {}
    for mapping_key, source_column in (column_mappings or {}).items():
        if not source_column:
            continue
        if "." not in mapping_key:
            raise ValueError(
                f"GradeRate mapping '{mapping_key}' must include its "
                "object name."
            )
        object_name, field_name = mapping_key.split(".", 1)
        if object_name.casefold() not in {"graderate", "graderatevalue"}:
            raise ValueError(
                f"Unknown GradeRate mapping object '{object_name}'."
            )
        if source_column not in dataframe.columns:
            raise ValueError(
                f"Mapped Excel column '{source_column}' for "
                f"'{mapping_key}' does not exist in the selected worksheet."
            )
        aliases = GRADE_RATE_SOURCE_ALIASES.get(field_name)
        if not aliases:
            continue
        target_column = aliases[0]
        prior_source = mapped_sources.get(target_column)
        if prior_source is not None and prior_source != source_column:
            raise ValueError(
                f"Multiple source columns were selected for GradeRate "
                f"field '{field_name}'."
            )
        mapped_sources[target_column] = source_column
        mapped_dataframe[target_column] = dataframe[source_column]
    return mapped_dataframe


def _run_grade_rate_pipeline(
    file_path,
    generation_overrides,
    settings_path,
    output_dir,
    reports_dir,
    column_mappings=None,
):
    source_system_owner = generation_overrides.get("source_system_owner")
    if not source_system_owner:
        clients = get_active_clients(settings_path)
        if len(clients) != 1:
            raise ValueError(
                "Expected exactly one active client for HDL source identifiers; "
                f"found {len(clients)}."
            )
        source_system_owner = clients[0].get("source_system_owner")
    if not source_system_owner:
        raise ValueError(
            "The active client has no source_system_owner configured."
        )

    dataframe = read_excel(file_path)
    dataframe = _mapped_grade_rate_source_dataframe(
        dataframe,
        column_mappings,
    )
    output_file = generate_grade_rate_dat(
        dataframe,
        source_system_owner,
        output_dir,
    )
    generated_rows = sum(
        line.startswith("MERGE|")
        for line in output_file.read_text(encoding="utf-8").splitlines()
    )
    if generated_rows == 0:
        raise ValueError(
            "No GradeRate records were generated. Verify the workbook "
            "contains GradeRate and GradeRateValue data."
        )
    return _builtin_generation_result(
        GRADE_RATE_TEMPLATE,
        output_file,
        len(dataframe),
        generated_rows,
        reports_dir,
        source_system_owner,
    )


def _mapped_grade_source_dataframe(dataframe, column_mappings):
    mapped_dataframe = _mapped_source_dataframe(dataframe, column_mappings)
    preferred_source_columns = {
        "EffectiveStartDate": "GradeEffectiveStartDate",
        "EffectiveEndDate": "GradeEffectiveEndDate",
    }
    for mapping_key, source_column in (column_mappings or {}).items():
        if not source_column:
            continue
        field_name = mapping_key.rsplit(".", 1)[-1]
        preferred_column = preferred_source_columns.get(field_name)
        if preferred_column:
            mapped_dataframe[preferred_column] = dataframe[source_column]
    return mapped_dataframe


def _run_grade_pipeline(
    file_path,
    generation_overrides,
    settings_path,
    output_dir,
    reports_dir,
    column_mappings=None,
):
    source_system_owner = generation_overrides.get("source_system_owner")
    if not source_system_owner:
        clients = get_active_clients(settings_path)
        if len(clients) != 1:
            raise ValueError(
                "Expected exactly one active client for HDL source identifiers; "
                f"found {len(clients)}."
            )
        source_system_owner = clients[0].get("source_system_owner")
    if not source_system_owner:
        raise ValueError(
            "The active client has no source_system_owner configured."
        )

    dataframe = read_excel(file_path)
    output_file = generate_grade_dat(
        _mapped_grade_source_dataframe(dataframe, column_mappings),
        source_system_owner,
        output_dir,
    )
    generated_rows = sum(
        line.startswith("MERGE|")
        for line in output_file.read_text(encoding="utf-8").splitlines()
    )
    if generated_rows == 0:
        raise ValueError(
            "No Grade records were generated. Verify the workbook contains "
            "grade code, grade name, and effective start date columns."
        )
    return _builtin_generation_result(
        GRADE_TEMPLATE,
        output_file,
        len(dataframe),
        generated_rows,
        reports_dir,
        source_system_owner,
    )


def _run_person_type_pipeline(
    file_path: str,
    generation_overrides: Dict[str, str],
    settings_path: str,
    output_dir: str,
    reports_dir: str,
    column_mappings=None,
) -> Dict[str, Any]:
    source_system_owner = generation_overrides.get("source_system_owner")
    if not source_system_owner:
        clients = get_active_clients(settings_path)
        if len(clients) != 1:
            raise ValueError(
                "Expected exactly one active client for HDL source identifiers; "
                f"found {len(clients)}."
            )
        source_system_owner = clients[0].get("source_system_owner")
    if not source_system_owner:
        raise ValueError(
            "The active client has no source_system_owner configured."
        )

    dataframe = read_excel(file_path)
    output_file = generate_person_type_dat(
        _mapped_source_dataframe(dataframe, column_mappings),
        source_system_owner,
        output_dir,
    )
    lines = output_file.read_text(encoding="utf-8").splitlines()
    generated_rows = sum(line.startswith("MERGE|") for line in lines)
    if generated_rows == 0:
        raise ValueError(
            "No PersonType records were generated. Verify the workbook "
            "contains PersonType definition columns."
        )

    report_dir = Path(reports_dir)
    report_dir.mkdir(parents=True, exist_ok=True)
    mapping_report_path = report_dir / "mapping_report.xlsx"
    validation_report_path = report_dir / "validation_errors.csv"
    pd.DataFrame(columns=[
        "object_name",
        "field_name",
        "excel_column",
        "mapping_type",
        "default_value",
        "status",
    ]).to_excel(mapping_report_path, index=False)
    pd.DataFrame(columns=[
        "ExcelRow",
        "Object",
        "Field",
        "Error",
    ]).to_csv(validation_report_path, index=False)

    output_path = output_file.as_posix()
    return {
        "success": True,
        "template_name": "PersonType",
        "output_file": output_path,
        "output_filename": output_file.name,
        "output_download_url": _output_download_url(output_file.name),
        "output_file_size_kb": round(output_file.stat().st_size / 1024, 2),
        "contact_output_file": None,
        "contact_output_filename": None,
        "contact_download_url": None,
        "person_disability_output_file": None,
        "person_disability_output_filename": None,
        "person_disability_download_url": None,
        "total_excel_rows": len(dataframe),
        "generated_rows": generated_rows,
        "skipped_rows": len(dataframe) - generated_rows,
        "validation_errors_count": 0,
        "validation_errors": [],
        "mapping_report_path": mapping_report_path.as_posix(),
        "validation_report_path": validation_report_path.as_posix(),
        "dat_preview": lines[:60],
        "generation_values": {
            "legislation_code": "",
            "legal_employer_name": "",
            "source_system_owner": source_system_owner,
        },
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def _run_legislative_data_group_pipeline(
    file_path: str,
    generation_overrides: Dict[str, str],
    settings_path: str,
    output_dir: str,
    reports_dir: str,
    column_mappings=None,
    generate_companion_outputs: bool = True,
) -> Dict[str, Any]:
    source_system_owner = generation_overrides.get("source_system_owner")
    if not source_system_owner:
        clients = get_active_clients(settings_path)
        if len(clients) != 1:
            raise ValueError(
                "Expected exactly one active client for HDL source identifiers; "
                f"found {len(clients)}."
            )
        source_system_owner = clients[0].get("source_system_owner")
    if not source_system_owner:
        raise ValueError(
            "The active client has no source_system_owner configured."
        )

    dataframe = read_excel(file_path)
    output_file = generate_legislative_data_group_dat(
        _mapped_source_dataframe(dataframe, column_mappings),
        source_system_owner,
        output_dir,
    )
    lines = output_file.read_text(encoding="utf-8").splitlines()
    generated_rows = sum(line.startswith("MERGE|") for line in lines)
    if generated_rows == 0:
        raise ValueError(
            "No LegislativeDataGroup records were generated. Verify the "
            "workbook contains LDG name, legislation code, currency code, "
            "and structure instance name values."
        )

    mapping_report_path = ""
    validation_report_path = ""
    if generate_companion_outputs:
        report_dir = Path(reports_dir)
        report_dir.mkdir(parents=True, exist_ok=True)
        mapping_report = report_dir / "mapping_report.xlsx"
        validation_report = report_dir / "validation_errors.csv"
        pd.DataFrame(columns=[
            "object_name",
            "field_name",
            "excel_column",
            "mapping_type",
            "default_value",
            "status",
        ]).to_excel(mapping_report, index=False)
        pd.DataFrame(columns=[
            "ExcelRow",
            "Object",
            "Field",
            "Error",
        ]).to_csv(validation_report, index=False)
        mapping_report_path = mapping_report.as_posix()
        validation_report_path = validation_report.as_posix()

    return {
        "success": True,
        "template_name": "LegislativeDataGroup",
        "output_file": output_file.as_posix(),
        "output_filename": output_file.name,
        "output_download_url": _output_download_url(output_file.name),
        "output_file_size_kb": round(output_file.stat().st_size, 2),
        "contact_output_file": None,
        "contact_output_filename": None,
        "contact_download_url": None,
        "person_disability_output_file": None,
        "person_disability_output_filename": None,
        "person_disability_download_url": None,
        "total_excel_rows": len(dataframe),
        "generated_rows": generated_rows,
        "skipped_rows": len(dataframe) - generated_rows,
        "validation_errors_count": 0,
        "validation_errors": [],
        "mapping_report_path": mapping_report_path,
        "validation_report_path": validation_report_path,
        "dat_preview": lines[:60],
        "generation_values": {
            "legislation_code": "",
            "legal_employer_name": "",
            "source_system_owner": source_system_owner,
        },
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def _run_legislative_data_group_translation_pipeline(
    file_path: str,
    generation_overrides: Dict[str, str],
    settings_path: str,
    output_dir: str,
    reports_dir: str,
    column_mappings=None,
    generate_companion_outputs: bool = True,
) -> Dict[str, Any]:
    source_system_owner = generation_overrides.get("source_system_owner")
    if not source_system_owner:
        clients = get_active_clients(settings_path)
        if len(clients) != 1:
            raise ValueError(
                "Expected exactly one active client for HDL source identifiers; "
                f"found {len(clients)}."
            )
        source_system_owner = clients[0].get("source_system_owner")
    if not source_system_owner:
        raise ValueError(
            "The active client has no source_system_owner configured."
        )

    dataframe = read_excel(file_path)
    output_file = generate_legislative_data_group_translation_dat(
        _mapped_source_dataframe(dataframe, column_mappings),
        source_system_owner,
        output_dir,
    )
    lines = output_file.read_text(encoding="utf-8").splitlines()
    generated_rows = sum(line.startswith("MERGE|") for line in lines)
    if generated_rows == 0:
        raise ValueError(
            "No LegislativeDataGroupTranslation records were generated. "
            "Verify the workbook contains LDG names and translated names."
        )

    mapping_report_path = ""
    validation_report_path = ""
    if generate_companion_outputs:
        report_dir = Path(reports_dir)
        report_dir.mkdir(parents=True, exist_ok=True)
        mapping_report = report_dir / "mapping_report.xlsx"
        validation_report = report_dir / "validation_errors.csv"
        pd.DataFrame(columns=[
            "object_name",
            "field_name",
            "excel_column",
            "mapping_type",
            "default_value",
            "status",
        ]).to_excel(mapping_report, index=False)
        pd.DataFrame(columns=[
            "ExcelRow",
            "Object",
            "Field",
            "Error",
        ]).to_csv(validation_report, index=False)
        mapping_report_path = mapping_report.as_posix()
        validation_report_path = validation_report.as_posix()

    return {
        "success": True,
        "template_name": "LegislativeDataGroupTranslation",
        "output_file": output_file.as_posix(),
        "output_filename": output_file.name,
        "output_download_url": _output_download_url(output_file.name),
        "output_file_size_kb": round(output_file.stat().st_size, 2),
        "contact_output_file": None,
        "contact_output_filename": None,
        "contact_download_url": None,
        "person_disability_output_file": None,
        "person_disability_output_filename": None,
        "person_disability_download_url": None,
        "total_excel_rows": len(dataframe),
        "generated_rows": generated_rows,
        "skipped_rows": len(dataframe) - generated_rows,
        "validation_errors_count": 0,
        "validation_errors": [],
        "mapping_report_path": mapping_report_path,
        "validation_report_path": validation_report_path,
        "dat_preview": lines[:60],
        "generation_values": {
            "legislation_code": "",
            "legal_employer_name": "",
            "source_system_owner": source_system_owner,
        },
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def _apply_batch_mapping_overrides(objects, column_mappings):
    fields_by_key = {
        f"{obj['object_name']}.{field['field_name']}": field
        for obj in objects
        for field in obj["fields"]
    }
    object_by_name = {
        obj["object_name"]: obj
        for obj in objects
    }
    for obj in objects:
        obj["mappings"] = [
            mapping
            for mapping in obj["mappings"]
            if not _is_internal_id_field(
                mapping.get("field_name") or ""
            )
        ]
    for mapping_key, source_column in column_mappings.items():
        if not source_column:
            continue
        if mapping_key not in fields_by_key:
            raise ValueError(
                f"Unknown batch mapping field '{mapping_key}'."
            )
        field = fields_by_key[mapping_key]
        field_name = field["field_name"]
        if _is_internal_id_field(field_name):
            raise ValueError(
                f"Internal ID field '{mapping_key}' cannot map to an Excel "
                "column; use the corresponding SourceSystemId reference."
            )
        if (
            field_name in {"SourceSystemId", "SourceSystemOwner"}
            or bool(field.get("generated_field"))
        ):
            raise ValueError(
                f"Generated field '{mapping_key}' cannot map to an Excel column."
            )
        obj_name = mapping_key.split(".", 1)[0]
        obj = object_by_name[obj_name]
        obj["mappings"] = [
            mapping
            for mapping in obj["mappings"]
            if mapping.get("field_name") != field_name
        ]
        obj["mappings"].append({
            "field_name": field_name,
            "excel_column": source_column,
            "mapping_type": "COLUMN",
            "required": field.get("required_flag"),
        })
    return objects


def run_hdl_generation_pipeline(
    file_path: str,
    template_name: str = "Worker",
    generation_overrides: Optional[Dict[str, str]] = None,
    settings_path: str = "config/settings.json",
    output_dir: str = "output",
    reports_dir: str = "reports",
    column_mappings: Optional[Dict[str, str]] = None,
    generate_companion_outputs: bool = True,
) -> Dict[str, Any]:
    """Execute the full HDL generation pipeline faithfully using backend logic."""
    generation_overrides = generation_overrides or {}

    if template_name.casefold() == "jobfamily":
        df = read_excel(file_path)
        df = _mapped_source_dataframe(df, column_mappings)
        output_file = generate_job_family(df, output_dir)
        lines = output_file.read_text(encoding="utf-8").splitlines()
        records_generated = sum(
            line.startswith("MERGE|") for line in lines
        )
        return {
            "success": True,
            "template_name": JOB_FAMILY_TEMPLATE["template_name"],
            "output_file": output_file.as_posix(),
            "output_filename": output_file.name,
            "output_download_url": _output_download_url(output_file.name),
            "output_file_size_kb": round(
                output_file.stat().st_size / 1024,
                2,
            ),
            "contact_output_file": None,
            "person_disability_output_file": None,
            "total_excel_rows": len(df),
            "generated_rows": records_generated,
            "skipped_rows": 0,
            "validation_errors_count": 0,
            "validation_errors": [],
            "mapping_report_path": "",
            "validation_report_path": "",
            "dat_preview": lines[:60],
            "generation_values": {
                "legislation_code": "",
                "legal_employer_name": "",
                "source_system_owner": "",
            },
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }

    if template_name.casefold() == "grade":
        return _run_grade_pipeline(
            file_path,
            generation_overrides,
            settings_path,
            output_dir,
            reports_dir,
            column_mappings,
        )

    if template_name.casefold() == "persondisability":
        return _run_person_disability_pipeline(
            file_path,
            generation_overrides,
            settings_path,
            output_dir,
            column_mappings,
        )

    if template_name.casefold() == "contact":
        return _run_contact_pipeline(
            file_path,
            generation_overrides,
            settings_path,
            output_dir,
            reports_dir,
            column_mappings,
        )

    if template_name.casefold() == "calendarevent":
        return _run_calendar_event_pipeline(
            file_path,
            output_dir,
            reports_dir,
            column_mappings,
        )

    if template_name.casefold() == "actionreasons":
        return _run_action_reasons_pipeline(
            file_path,
            generation_overrides,
            settings_path,
            output_dir,
            reports_dir,
            column_mappings,
        )

    if template_name.casefold() == "actions":
        return _run_actions_pipeline(
            file_path,
            generation_overrides,
            settings_path,
            output_dir,
            reports_dir,
            column_mappings,
        )

    if template_name.casefold() == "graderate":
        return _run_grade_rate_pipeline(
            file_path,
            generation_overrides,
            settings_path,
            output_dir,
            reports_dir,
            column_mappings,
        )

    if template_name.casefold() == "actionreasonstranslation":
        return _run_action_reasons_translation_pipeline(
            file_path,
            output_dir,
            reports_dir,
            column_mappings,
        )

    if template_name.casefold() == "actionstranslation":
        return _run_actions_translation_pipeline(
            file_path,
            output_dir,
            reports_dir,
            column_mappings,
        )

    if template_name.casefold() == "resourceexception":
        return _run_resource_exception_pipeline(
            file_path,
            generation_overrides,
            settings_path,
            output_dir,
            reports_dir,
            column_mappings,
        )

    if template_name.casefold() == "persontype":
        return _run_person_type_pipeline(
            file_path,
            generation_overrides,
            settings_path,
            output_dir,
            reports_dir,
            column_mappings,
        )

    if template_name.casefold() == "legislativedatagroup":
        return _run_legislative_data_group_pipeline(
            file_path,
            generation_overrides,
            settings_path,
            output_dir,
            reports_dir,
            column_mappings,
            generate_companion_outputs,
        )

    if template_name.casefold() == "legislativedatagrouptranslation":
        return _run_legislative_data_group_translation_pipeline(
            file_path,
            generation_overrides,
            settings_path,
            output_dir,
            reports_dir,
            column_mappings,
            generate_companion_outputs,
        )

    template, objects = load_template_configuration(settings_path, template_name)
    objects = _apply_batch_mapping_overrides(
        objects,
        column_mappings or {},
    )
    objects = backend_main._worker_template_objects(
        template.get("template_name") or "",
        objects,
    )
    df = read_excel(file_path)
    alias_map = load_field_mappings()
    canonical_rows, _ = map_excel_dataframe(df, alias_map)
    canonical_rows = _apply_canonical_column_mappings(
        df,
        canonical_rows,
        column_mappings,
    )

    generation_defaults = backend_main._worker_generation_defaults(
        canonical_rows,
        template["source_system_owner"],
    )

    # Merge user overrides with auto-detected defaults
    generation_values = {
        "legislation_code": backend_main._normalize_legislation_code(
            generation_overrides.get("legislation_code")
            or generation_defaults.get("legislation_code")
            or ""
        ),
        "legal_employer_name": generation_overrides.get("legal_employer_name")
        or generation_defaults.get("legal_employer_name")
        or "",
        "source_system_owner": generation_overrides.get("source_system_owner")
        or generation_defaults.get("source_system_owner")
        or template.get("source_system_owner")
        or "",
    }

    identifier_config = load_national_identifier_config()
    alias_index = build_alias_index(alias_map)

    normalized_columns = {}
    for column in df.columns:
        normalized = normalize_column_name(column)
        if normalized in normalized_columns:
            raise ValueError(
                f"Excel columns '{normalized_columns[normalized]}' and "
                f"'{column}' normalize to the same field name."
            )
        normalized_columns[normalized] = column

    mapping_lookup, mapping_report = backend_main._make_mapping_lookup(
        objects,
        normalized_columns,
        alias_index,
    )

    output_objects = [
        {
            "object_name": obj["object_name"],
            "metadata_fields": (
                [field["field_name"] for field in obj["fields"]]
                if (
                    template["template_name"].casefold() == "worker"
                    and obj["object_name"] == "PersonEmail"
                )
                else backend_main.WORKER_METADATA_FIELDS.get(
                    obj["object_name"],
                    [field["field_name"] for field in obj["fields"]],
                ) if template["template_name"].casefold() == "worker" else [
                    field["field_name"] for field in obj["fields"]
                ]
            ),
            "records": [],
        }
        for obj in objects
    ]
    validation_errors = []

    for data_index, (_, row) in enumerate(df.iterrows()):
        row_index = data_index + 2
        email_entries = backend_main._email_entries(
            row,
            normalized_columns,
            alias_index,
        )
        mapped_row_objects = []
        identifier_row = canonical_rows[data_index]
        country_field = identifier_config.get("country_field")
        if is_empty(identifier_row.get(country_field)):
            for country_source in ("Country", "LegislationCode"):
                country_value = identifier_row.get(country_source)
                if not is_empty(country_value):
                    identifier_row = {
                        **identifier_row,
                        country_field: country_value,
                    }
                    break
        row_errors = [
            {
                "ExcelRow": row_index,
                "Object": "",
                "Field": "NationalIdentifierNumber",
                "Error": error,
            }
            for error in validate_national_identifier(
                identifier_row,
                identifier_config,
            )
        ]

        for obj in objects:
            record = {}
            for field in obj["fields"]:
                field_name = field["field_name"]
                mapping = mapping_lookup[(obj["object_name"], field_name)]
                if obj["object_name"] == "PersonEmail" and field_name in {
                    "EmailAddress", "EmailType", "PrimaryFlag"
                }:
                    value = (
                        email_entries[0].get(field_name, "")
                        if email_entries else ""
                    )
                elif mapping and mapping.get("mapping_type") == "GENERATED":
                    value = backend_main._generated_value(
                        field_name,
                        obj["object_name"],
                        canonical_rows[data_index].get("PersonNumber", ""),
                        canonical_rows[data_index],
                        generation_values["source_system_owner"],
                        template["template_name"],
                        record_index=data_index + 1,
                    )
                else:
                    value = backend_main._map_value(mapping, row, normalized_columns)
                if normalize_column_name(field_name) == "nationalidentifiernumber":
                    value = normalize_national_identifier(value)
                record[field_name] = value

                is_required = backend_main._is_required(field.get("required_flag"))
                output_fields = backend_main.WORKER_METADATA_FIELDS.get(
                    obj["object_name"]
                )
                if (
                    output_fields is not None
                    and field_name not in output_fields
                ):
                    is_required = False
                if mapping and backend_main._is_required(mapping.get("required")):
                    is_required = True

                identifier_may_be_blank = (
                    identifier_config["allow_blank"]
                    and normalize_column_name(field_name)
                    == "nationalidentifiernumber"
                )
                if is_required and not identifier_may_be_blank and (
                    value is None or pd.isna(value) or str(value).strip() == ""
                ):
                    row_errors.append({
                        "ExcelRow": row_index,
                        "Object": obj["object_name"],
                        "Field": field_name,
                        "Error": "Required value is blank.",
                    })

            mapped_row_objects.append(record)

        if row_errors:
            validation_errors.extend(row_errors)
            continue

        source_columns = {
            normalize_column_name(column): value
            for column, value in row.items()
        }
        source_columns["Nationality"] = canonical_rows[data_index].get(
            "Nationality",
            "",
        )
        person_number = canonical_rows[data_index].get("PersonNumber", "")
        for output_obj, record in zip(output_objects, mapped_row_objects):
            if template["template_name"].casefold() == "worker":
                if output_obj["object_name"] == "PersonEmail":
                    for email_entry in email_entries:
                        email_record = {
                            **record,
                            **email_entry,
                        }
                        base_source_id = email_record.get(
                            "SourceSystemId",
                            "",
                        )
                        if base_source_id and len(email_entries) > 1:
                            email_record["SourceSystemId"] = (
                                f"{base_source_id}_"
                                f"{email_entry['EmailType'].casefold()}"
                            )
                        output_obj["records"].append(
                            backend_main._build_worker_metadata_record(
                                output_obj["object_name"],
                                email_record,
                                person_number,
                                generation_values["source_system_owner"],
                                source_columns,
                                template["template_name"],
                                generation_values,
                                output_obj["metadata_fields"],
                            )
                        )
                    continue
                record = backend_main._build_worker_metadata_record(
                    output_obj["object_name"],
                    record,
                    person_number,
                    generation_values["source_system_owner"],
                    source_columns,
                    template["template_name"],
                    generation_values,
                )
            output_obj["records"].append(record)

    if template["template_name"].casefold() == "job":
        for output_obj in output_objects:
            if output_obj["object_name"] == "Job":
                output_obj["records"] = _normalize_job_records(
                    output_obj["records"],
                    generation_values["source_system_owner"],
                )

    dat_name = Path(
        template.get("output_file_name")
        or template.get("dat_file")
        or ""
    ).name
    if not dat_name:
        dat_name = f"{template['template_name'].replace(' ', '_')}.dat"
    if Path(dat_name).suffix.lower() != ".dat":
        raise ValueError(
            f"Template output file must have a .dat extension: {dat_name}"
        )

    out_dir_path = Path(output_dir)
    rep_dir_path = Path(reports_dir)
    out_dir_path.mkdir(parents=True, exist_ok=True)
    rep_dir_path.mkdir(parents=True, exist_ok=True)

    output_file = out_dir_path / dat_name
    disability_output_file = None
    disability_validation_errors = []
    if (
        generate_companion_outputs
        and template["template_name"].casefold() == "worker"
    ):
        disability_lookups = (
            get_person_disability_lookups(settings_path)
            if has_disability_data(canonical_rows)
            else []
        )
        try:
            disability_output_file = generate_person_disability(
                canonical_rows,
                disability_lookups,
                out_dir_path,
                legislation_code=generation_values["legislation_code"],
            )
        except ValueError as exc:
            error = _person_disability_validation_error(exc)
            if error is None:
                raise
            disability_validation_errors.append(error)
            (out_dir_path / "PersonDisability.dat").unlink(missing_ok=True)
    generate_hdl_bundle(output_file, output_objects)
    contact_output_file = None
    if template["template_name"].casefold() == "worker":
        invalid_rows = {
            error["ExcelRow"] - 2
            for error in validation_errors
            if "ExcelRow" in error
        }
        valid_rows = df.iloc[
            [index for index in range(len(df)) if index not in invalid_rows]
        ]
        contact_output_file = generate_contact_dat(
            valid_rows,
            generation_values["source_system_owner"],
            out_dir_path,
        )

    mapping_report_path = rep_dir_path / "mapping_report.xlsx"
    validation_report_path = rep_dir_path / "validation_errors.csv"
    validation_errors.extend(disability_validation_errors)
    pd.DataFrame(mapping_report).to_excel(mapping_report_path, index=False)
    pd.DataFrame(validation_errors).to_csv(validation_report_path, index=False)

    generated_rows = len(output_objects[0]["records"]) if output_objects else 0
    skipped_rows = len({
        item["ExcelRow"]
        for item in validation_errors
        if item.get("Object") != "PersonDisability"
    })

    # Read preview of DAT file
    dat_snippet = []
    if output_file.exists():
        with output_file.open("r", encoding="utf-8") as f:
            for _ in range(60):
                line = f.readline()
                if not line:
                    break
                dat_snippet.append(line.rstrip("\r\n"))

    return {
        "success": True,
        "template_name": template["template_name"],
        "output_file": str(output_file.as_posix()),
        "output_filename": dat_name,
        "output_download_url": _output_download_url(dat_name),
        "output_file_size_kb": round(output_file.stat().st_size / 1024, 2),
        "contact_output_file": (
            str(contact_output_file.as_posix())
            if contact_output_file else None
        ),
        "contact_output_filename": (
            contact_output_file.name if contact_output_file else None
        ),
        "contact_download_url": _output_download_url(
            contact_output_file.name if contact_output_file else None
        ),
        "person_disability_output_file": (
            str(disability_output_file.as_posix())
            if disability_output_file else None
        ),
        "person_disability_output_filename": (
            disability_output_file.name if disability_output_file else None
        ),
        "person_disability_download_url": _output_download_url(
            disability_output_file.name if disability_output_file else None
        ),
        "total_excel_rows": len(df),
        "generated_rows": generated_rows,
        "skipped_rows": skipped_rows,
        "validation_errors_count": len(validation_errors),
        "validation_errors": validation_errors[:100],  # Return up to first 100 in payload
        "mapping_report_path": str(mapping_report_path.as_posix()),
        "validation_report_path": str(validation_report_path.as_posix()),
        "dat_preview": dat_snippet,
        "generation_values": generation_values,
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def _batch_mapping_issues(dataframe, template_name, column_mappings, settings_path):
    schema = get_batch_template_schema(template_name, settings_path)
    fields = schema["source_fields"]
    fields_by_key = {
        field["mapping_key"]: field
        for field in fields
    }
    mapping_issues = []
    selected_columns = {}

    for mapping_key, source_column in (column_mappings or {}).items():
        if mapping_key not in fields_by_key:
            mapping_issues.append(f"Unknown mapping field '{mapping_key}'.")
            continue
        if not source_column:
            continue
        if source_column not in dataframe.columns:
            mapping_issues.append(
                f"Selected Excel column '{source_column}' for "
                f"'{mapping_key}' is not in the uploaded worksheet."
            )
            continue
        selected_columns[mapping_key] = source_column

    unsupported_required = [
        field["mapping_key"]
        for field in fields
        if field.get("internal_id") and field.get("target_required")
    ]
    mapping_issues.extend(
        f"Required field '{field}' is an internal ID and cannot be mapped; "
        "use a SourceSystemId reference."
        for field in unsupported_required
    )

    missing_required = [
        field["mapping_key"]
        for field in fields
        if field["required"]
        and field["mapping_key"] not in selected_columns
    ]
    if normalize_column_name(template_name) == "contact":
        contact_fields = {"FatherName", "MotherName", "Spouse"}
        if not any(field in selected_columns for field in contact_fields):
            missing_required.append(
                "At least one of FatherName, MotherName, or Spouse"
            )

    return missing_required, mapping_issues


def _deduplicate_hdl_dat_file(output_file):
    lines = output_file.read_text(encoding="utf-8").splitlines()
    metadata_fields = {}
    seen_records = {}
    deduplicated = []
    removed_count = 0

    for line in lines:
        parts = line.split("|")
        if parts[0] == "METADATA" and len(parts) >= 3:
            metadata_fields[parts[1]] = parts[2:]
            deduplicated.append(line)
            continue
        if parts[0] != "MERGE" or len(parts) < 3:
            deduplicated.append(line)
            continue

        object_name = parts[1]
        fields = metadata_fields.get(object_name, [])
        values = parts[2:]
        key = tuple(
            (field, values[index] if index < len(values) else "")
            for index, field in enumerate(fields)
            if field != "SourceSystemId"
        )
        object_seen = seen_records.setdefault(object_name, set())
        if key in object_seen:
            removed_count += 1
            continue
        object_seen.add(key)
        deduplicated.append(line)

    if removed_count:
        output_file.write_text(
            "\n".join(deduplicated) + "\n",
            encoding="utf-8",
        )
    return removed_count


def validate_batch_generation(
    file_path: str,
    templates: List[Dict[str, Any]],
    generation_overrides: Optional[Dict[str, str]] = None,
    settings_path: str = "config/settings.json",
) -> Dict[str, Any]:
    """Dry-run selected templates and return per-template validation results."""
    if not templates:
        raise ValueError("Select at least one HDL template.")

    dataframe = read_excel(file_path)
    generation_overrides = generation_overrides or {}
    results = []
    seen_names = set()

    for selection in sorted(
        templates,
        key=lambda item: _batch_load_order(item["template_name"]),
    ):
        template_name = selection["template_name"]
        name_key = template_name.casefold()
        if name_key in seen_names:
            raise ValueError(
                f"Template '{template_name}' was selected more than once."
            )
        seen_names.add(name_key)
        column_mappings = selection.get("column_mappings") or {}
        missing_required, mapping_issues = _batch_mapping_issues(
            dataframe,
            template_name,
            column_mappings,
            settings_path,
        )
        errors = list(mapping_issues)
        generated_rows = 0

        if missing_required:
            errors.extend(
                f"Required field has no selected Excel column: {field}"
                for field in missing_required
            )
        if not errors:
            try:
                with tempfile.TemporaryDirectory() as temporary_directory:
                    output_dir = Path(temporary_directory) / "output"
                    reports_dir = Path(temporary_directory) / "reports"
                    generated = run_hdl_generation_pipeline(
                        file_path,
                        template_name=template_name,
                        generation_overrides=generation_overrides,
                        settings_path=settings_path,
                        output_dir=str(output_dir),
                        reports_dir=str(reports_dir),
                        column_mappings=column_mappings,
                        generate_companion_outputs=False,
                    )
                    generated_rows = generated["generated_rows"]
                    generated_file = Path(generated["output_file"])
                    if generated_file.is_file():
                        generated_rows = sum(
                            line.startswith("MERGE|")
                            for line in generated_file.read_text(
                                encoding="utf-8",
                            ).splitlines()
                        )
                    errors.extend(generated.get("validation_errors", []))
            except Exception as exc:
                errors.append(str(exc))

        results.append({
            "template_name": template_name,
            "category": _batch_template_category(template_name),
            "row_count": len(dataframe),
            "generated_rows": generated_rows,
            "missing_required_columns": missing_required,
            "errors": errors,
            "ready": (
                not missing_required
                and not mapping_issues
                and not errors
                and generated_rows > 0
            ),
        })

    return {
        "file_path": file_path,
        "row_count": len(dataframe),
        "templates": results,
        "can_generate": all(result["ready"] for result in results),
    }


def generate_batch_hdl(
    file_path: str,
    templates: List[Dict[str, Any]],
    generation_overrides: Optional[Dict[str, str]] = None,
    settings_path: str = "config/settings.json",
    output_dir: str = "output",
) -> Dict[str, Any]:
    """Generate selected template DAT files in load order and package a ZIP."""
    validation = validate_batch_generation(
        file_path,
        templates,
        generation_overrides,
        settings_path,
    )
    if not validation["can_generate"]:
        raise ValueError(
            "Batch validation failed. Review missing required mappings and "
            "generation errors before generating."
        )

    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    generation_overrides = generation_overrides or {}
    generated_files = []
    template_results = []
    filenames = set()

    with tempfile.TemporaryDirectory() as temporary_directory:
        temporary_root = Path(temporary_directory)
        staging_dir = temporary_root / "staging"
        staging_dir.mkdir()
        for index, selection in enumerate(sorted(
            templates,
            key=lambda item: _batch_load_order(item["template_name"]),
        )):
            template_name = selection["template_name"]
            template_dir = temporary_root / f"template_{index}"
            template_output_dir = template_dir / "output"
            reports_dir = template_dir / "reports"
            result = run_hdl_generation_pipeline(
                file_path,
                template_name=template_name,
                generation_overrides=generation_overrides,
                settings_path=settings_path,
                output_dir=str(template_output_dir),
                reports_dir=str(reports_dir),
                column_mappings=selection.get("column_mappings") or {},
                generate_companion_outputs=False,
            )
            dat_file = template_output_dir / result["output_filename"]
            if not dat_file.is_file():
                raise ValueError(
                    f"{template_name} generation did not create its expected "
                    f"output file '{result['output_filename']}'."
                )
            if _batch_template_category(template_name) != "Employee Data":
                result["deduplicated_rows"] = _deduplicate_hdl_dat_file(
                    dat_file
                )
            else:
                result["deduplicated_rows"] = 0

            filename = dat_file.name.casefold()
            if filename in filenames:
                raise ValueError(
                    f"Multiple selected templates generate the same output "
                    f"file '{dat_file.name}'."
                )
            filenames.add(filename)
            staged_file = staging_dir / dat_file.name
            shutil.copy2(dat_file, staged_file)
            generated_files.append(staged_file)
            result["generated_rows"] = sum(
                line.startswith("MERGE|")
                for line in staged_file.read_text(
                    encoding="utf-8",
                ).splitlines()
            )
            template_results.append({
                "template_name": template_name,
                "output_filename": staged_file.name,
                "generated_rows": result["generated_rows"],
                "deduplicated_rows": result["deduplicated_rows"],
                "validation_errors_count": result.get(
                    "validation_errors_count",
                    0,
                ),
                "validation_errors": result.get("validation_errors", []),
            })

        zip_filename = (
            "HDL_Bundle_"
            + datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            + ".zip"
        )
        staged_zip = staging_dir / zip_filename
        with zipfile.ZipFile(
            staged_zip,
            mode="w",
            compression=zipfile.ZIP_DEFLATED,
        ) as archive:
            for staged_file in generated_files:
                archive.write(staged_file, arcname=staged_file.name)

        for staged_file in generated_files:
            os.replace(staged_file, output_path / staged_file.name)
        os.replace(staged_zip, output_path / zip_filename)

    return {
        "success": True,
        "zip_filename": zip_filename,
        "zip_download_url": _output_download_url(zip_filename),
        "templates": template_results,
        "output_filenames": [item.name for item in generated_files],
    }


def list_reports_and_outputs(output_dir: str = "output", reports_dir: str = "reports") -> Dict[str, Any]:
    """Retrieve all generated outputs and reports for management and downloading."""
    out_path = Path(output_dir)
    rep_path = Path(reports_dir)

    outputs = []
    if out_path.exists():
        for f in out_path.glob("*.dat"):
            stat = f.stat()
            outputs.append({
                "filename": f.name,
                "type": "HDL DAT File",
                "size_kb": round(stat.st_size / 1024, 2),
                "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
                "path": f"output/{f.name}",
            })
        for f in out_path.glob("*.zip"):
            stat = f.stat()
            outputs.append({
                "filename": f.name,
                "type": "HDL ZIP Bundle",
                "size_kb": round(stat.st_size / 1024, 2),
                "modified": datetime.fromtimestamp(stat.st_mtime).strftime(
                    "%Y-%m-%d %H:%M:%S"
                ),
                "path": f"output/{f.name}",
            })

    reports = []
    if rep_path.exists():
        for f in rep_path.glob("*.*"):
            if f.suffix.lower() in (".xlsx", ".csv", ".json"):
                stat = f.stat()
                reports.append({
                    "filename": f.name,
                    "type": "Excel Report" if f.suffix.lower() == ".xlsx" else "CSV Report",
                    "size_kb": round(stat.st_size / 1024, 2),
                    "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
                    "path": f"reports/{f.name}",
                })

    outputs.sort(key=lambda x: x["modified"], reverse=True)
    reports.sort(key=lambda x: x["modified"], reverse=True)
    return {"outputs": outputs, "reports": reports}


def read_full_dat_file(filename: str, output_dir: str = "output", max_lines: int = 500) -> Dict[str, Any]:
    """Read contents of an HDL DAT file for interactive browser inspection."""
    file_path = Path(output_dir) / filename
    if not file_path.exists():
        raise FileNotFoundError(f"File {filename} not found in {output_dir}")

    lines = []
    metadata_count = 0
    merge_count = 0
    with file_path.open("r", encoding="utf-8") as f:
        for idx, line in enumerate(f):
            clean_line = line.rstrip("\r\n")
            if clean_line.startswith("METADATA|"):
                metadata_count += 1
            elif clean_line.startswith("MERGE|"):
                merge_count += 1
            if idx < max_lines:
                lines.append(clean_line)

    return {
        "filename": filename,
        "total_lines_read": len(lines),
        "metadata_lines": metadata_count,
        "merge_lines": merge_count,
        "has_more": (metadata_count + merge_count) > max_lines,
        "lines": lines,
    }


def read_report_csv(
    filename: str,
    reports_dir: str = "reports",
    max_rows: int = 150,
) -> Dict[str, Any]:
    """Read a CSV report (e.g. validation_errors.csv) for interactive preview."""
    file_path = Path(reports_dir) / filename
    if not file_path.exists():
        raise FileNotFoundError(f"Report file {filename} not found in {reports_dir}")

    if file_path.stat().st_size == 0:
        return {
            "filename": filename,
            "total_rows": 0,
            "columns": ["ExcelRow", "Object", "Field", "Error"],
            "rows": [],
        }

    try:
        df = pd.read_csv(file_path)
    except pd.errors.EmptyDataError:
        return {
            "filename": filename,
            "total_rows": 0,
            "columns": ["ExcelRow", "Object", "Field", "Error"],
            "rows": [],
        }
    except Exception as exc:
        raise ValueError(f"Failed to read CSV report: {exc}")

    total_rows = len(df)
    columns = [str(c) for c in df.columns]
    rows = df.head(max_rows).fillna("").to_dict(orient="records")

    return {
        "filename": filename,
        "total_rows": total_rows,
        "columns": columns,
        "rows": rows,
    }


def read_report_excel(
    filename: str,
    reports_dir: str = "reports",
    max_rows: int = 250,
) -> Dict[str, Any]:
    """Read an Excel report (e.g. mapping_report.xlsx) for interactive preview."""
    file_path = Path(reports_dir) / filename
    if not file_path.exists():
        raise FileNotFoundError(f"Report file {filename} not found in {reports_dir}")

    try:
        df = pd.read_excel(file_path)
    except Exception as exc:
        raise ValueError(f"Failed to read Excel report: {exc}")

    total_rows = len(df)
    columns = [str(c) for c in df.columns]
    rows = df.head(max_rows).fillna("").to_dict(orient="records")

    return {
        "filename": filename,
        "total_rows": total_rows,
        "columns": columns,
        "rows": rows,
    }
