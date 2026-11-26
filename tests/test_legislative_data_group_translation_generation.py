from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import pandas as pd

from frontend.service import (
    analyze_field_mappings,
    get_available_templates,
    run_hdl_generation_pipeline,
)
from src.legislative_data_group_translation_generator import (
    build_legislative_data_group_translation_output_objects,
)


class LegislativeDataGroupTranslationGenerationTests(unittest.TestCase):
    def setUp(self):
        self.dataframe = pd.DataFrame([
            {
                "Name": "Republic Day",
                "Language": "AR",
                "LegislativeDataGroupName": "IN Legislative Data Group",
                "LegislationCode": "IN",
                "DefaultCurrencyCode": "INR",
                "StructureInstanceName": "MIS Costing",
                "LegislativeDataGroupTranslatedName": (
                    "مجموعة البيانات التشريعية IN"
                ),
            },
            {
                "Name": "UAE National Day",
                "Language": "AR",
                "LegislativeDataGroupName": "AE Legislative Data Group",
                "LegislationCode": "AE",
                "DefaultCurrencyCode": "AED",
                "StructureInstanceName": "MIS Costing",
                "LegislativeDataGroupTranslatedName": (
                    "مجموعة البيانات التشريعية AE"
                ),
            },
            {
                "Name": "Repeat UAE row",
                "Language": "AR",
                "LegislativeDataGroupName": "AE Legislative Data Group",
                "LegislationCode": "AE",
                "DefaultCurrencyCode": "AED",
                "StructureInstanceName": "MIS Costing",
                "LegislativeDataGroupTranslatedName": (
                    "مجموعة البيانات التشريعية AE"
                ),
            },
            {
                "Name": "Qatar National Day",
                "Language": "AR",
                "LegislativeDataGroupName": "QA Legislative Data Group",
                "LegislationCode": "QA",
                "DefaultCurrencyCode": "QAR",
                "StructureInstanceName": "MIS Costing",
                "LegislativeDataGroupTranslatedName": (
                    "مجموعة البيانات التشريعية QA"
                ),
            },
        ])

    def test_builds_distinct_translations_with_matching_ldg_source_ids(self):
        output_objects = build_legislative_data_group_translation_output_objects(
            self.dataframe,
            "GRISMO",
        )

        records = output_objects[0]["records"]
        self.assertEqual(len(records), 3)
        self.assertEqual(records[0], {
            "Language": "AR",
            "LegislativeDataGroupId": "",
            "Name": "مجموعة البيانات التشريعية IN",
            "SourceSystemOwner": "GRISMO",
            "SourceSystemId": "LDG_IN_01",
            "BaseName": "IN Legislative Data Group",
        })
        self.assertEqual(records[-1]["SourceSystemId"], "LDG_QA_01")

    def test_reports_missing_translation_fields(self):
        with self.assertRaisesRegex(ValueError, "requires Excel columns"):
            build_legislative_data_group_translation_output_objects(
                pd.DataFrame({"Language": ["AR"]}),
                "GRISMO",
            )

    def test_registers_template_and_maps_translation_aliases(self):
        with patch("frontend.service.get_templates", return_value=[]):
            templates = get_available_templates()

        template = next(
            item for item in templates
            if item["template_name"] == "LegislativeDataGroupTranslation"
        )
        self.assertEqual(
            template["output_file_name"],
            "LegislativeDataGroupTranslation.dat",
        )
        self.assertEqual(template["category"], "Organization Data")

        with patch("frontend.service.read_excel", return_value=self.dataframe):
            mappings = analyze_field_mappings(
                "input.xlsx",
                "LegislativeDataGroupTranslation",
            )["mappings"]

        mapping_by_field = {
            item["field_name"]: item for item in mappings
        }
        self.assertEqual(
            mapping_by_field["Name"]["excel_column"],
            "LegislativeDataGroupTranslatedName",
        )
        self.assertEqual(
            mapping_by_field["BaseName"]["excel_column"],
            "LegislativeDataGroupName",
        )

    def test_pipeline_writes_translation_dat(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output_dir = Path(temp_dir) / "output"
            reports_dir = Path(temp_dir) / "reports"
            with (
                patch(
                    "frontend.service.get_active_clients",
                    return_value=[{"source_system_owner": "GRISMO"}],
                ),
                patch(
                    "frontend.service.read_excel",
                    return_value=self.dataframe,
                ),
            ):
                result = run_hdl_generation_pipeline(
                    "input.xlsx",
                    template_name="LegislativeDataGroupTranslation",
                    output_dir=str(output_dir),
                    reports_dir=str(reports_dir),
                    generate_companion_outputs=False,
                )

            content = Path(result["output_file"]).read_text(
                encoding="utf-8",
            )
            self.assertIn(
                "METADATA|LegislativeDataGroupTranslation|Language|"
                "LegislativeDataGroupId|Name|SourceSystemOwner|"
                "SourceSystemId|BaseName",
                content,
            )
            self.assertIn(
                "MERGE|LegislativeDataGroupTranslation|AR||"
                "مجموعة البيانات التشريعية QA|GRISMO|LDG_QA_01|"
                "QA Legislative Data Group",
                content,
            )
            self.assertEqual(result["generated_rows"], 3)
            self.assertFalse(reports_dir.exists())


if __name__ == "__main__":
    unittest.main()
