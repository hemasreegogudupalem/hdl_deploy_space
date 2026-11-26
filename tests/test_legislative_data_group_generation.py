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
from src.legislative_data_group_generator import (
    build_legislative_data_group_output_objects,
)


class LegislativeDataGroupGenerationTests(unittest.TestCase):
    def setUp(self):
        self.dataframe = pd.DataFrame([
            {
                "Name": "Republic Day",
                "LegislativeDataGroupName": "IN Legislative Data Group",
                "LegislationCode": "IN",
                "DefaultCurrencyCode": "INR",
                "StructureInstanceName": "MIS Costing",
            },
            {
                "Name": "Saudi Founding Day",
                "LegislativeDataGroupName": "SA Legislative Data Group",
                "LegislationCode": "SA",
                "DefaultCurrencyCode": "SAR",
                "StructureInstanceName": "MIS Costing",
            },
            {
                "Name": "UAE National Day",
                "LegislativeDataGroupName": "AE Legislative Data Group",
                "LegislationCode": "AE",
                "DefaultCurrencyCode": "AED",
                "StructureInstanceName": "MIS Costing",
            },
            {
                "Name": "Qatar National Day",
                "LegislativeDataGroupName": "QA Legislative Data Group",
                "LegislationCode": "QA",
                "DefaultCurrencyCode": "QAR",
                "StructureInstanceName": "MIS Costing",
            },
            {
                "Name": "Another India event",
                "LegislativeDataGroupName": "IN Legislative Data Group",
                "LegislationCode": "IN",
                "DefaultCurrencyCode": "INR",
                "StructureInstanceName": "MIS Costing",
            },
        ])

    def test_builds_distinct_ldgs_with_generated_ids(self):
        output_objects = build_legislative_data_group_output_objects(
            self.dataframe,
            "GRISMO",
        )

        self.assertEqual(len(output_objects[0]["records"]), 4)
        self.assertEqual(
            [record["SourceSystemId"] for record in output_objects[0]["records"]],
            ["LDG_IN_01", "LDG_SA_01", "LDG_AE_01", "LDG_QA_01"],
        )
        self.assertEqual(
            output_objects[0]["records"][0],
            {
                "Name": "IN Legislative Data Group",
                "CostAllocationIdFlexNumber": "",
                "LegislationCode": "IN",
                "LegislativeDataGroupId": "",
                "DefaultCurrencyCode": "INR",
                "DefaultCurrency": "",
                "StructureInstanceName": "MIS Costing",
                "Territory": "",
                "SourceSystemOwner": "GRISMO",
                "SourceSystemId": "LDG_IN_01",
            },
        )

    def test_reports_missing_columns_and_incomplete_rows(self):
        with self.assertRaisesRegex(ValueError, "requires Excel columns"):
            build_legislative_data_group_output_objects(
                pd.DataFrame({"Name": ["QA Legislative Data Group"]}),
                "GRISMO",
            )

        incomplete = self.dataframe.iloc[:1].copy()
        incomplete.loc[0, "DefaultCurrencyCode"] = None
        with self.assertRaisesRegex(ValueError, "Excel row 2"):
            build_legislative_data_group_output_objects(incomplete, "GRISMO")

    def test_registers_builtin_template_and_maps_workbook_name_alias(self):
        with patch("frontend.service.get_templates", return_value=[]):
            templates = get_available_templates()

        template = next(
            item for item in templates
            if item["template_name"] == "LegislativeDataGroup"
        )
        self.assertEqual(template["output_file_name"], "LegislativeDataGroup.dat")
        self.assertEqual(template["category"], "Organization Data")

        with patch("frontend.service.read_excel", return_value=self.dataframe):
            mappings = analyze_field_mappings(
                "input.xlsx",
                "LegislativeDataGroup",
            )["mappings"]

        name_mapping = next(item for item in mappings if item["field_name"] == "Name")
        self.assertEqual(
            name_mapping["excel_column"],
            "LegislativeDataGroupName",
        )
        self.assertEqual(name_mapping["status"], "MAPPED")

    def test_pipeline_writes_legislative_data_group_dat(self):
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
                    template_name="LegislativeDataGroup",
                    output_dir=str(output_dir),
                    reports_dir=str(reports_dir),
                )

            content = Path(result["output_file"]).read_text(
                encoding="utf-8",
            )
            self.assertIn(
                "METADATA|LegislativeDataGroup|Name|"
                "CostAllocationIdFlexNumber|LegislationCode|"
                "LegislativeDataGroupId|DefaultCurrencyCode|DefaultCurrency|"
                "StructureInstanceName|Territory|SourceSystemOwner|SourceSystemId",
                content,
            )
            self.assertIn(
                "MERGE|LegislativeDataGroup|IN Legislative Data Group||"
                "IN||INR||MIS Costing||GRISMO|LDG_IN_01",
                content,
            )
            self.assertEqual(result["generated_rows"], 4)
            self.assertTrue(Path(result["mapping_report_path"]).is_file())
            self.assertTrue(Path(result["validation_report_path"]).is_file())


if __name__ == "__main__":
    unittest.main()
