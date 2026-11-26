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
from src.actions_generator import build_actions_output_objects


class ActionsGenerationTests(unittest.TestCase):
    def setUp(self):
        self.dataframe = pd.DataFrame([
            {
                "ActionCode": "SALARY_REVISION",
                "ActionName": "Employee Salary Revision",
                "ActionDescription": "Revision of employee salary",
                "ActionTypeCode": "EMPL_CHANGE",
                "ActionStartDate": "2026-11-16",
                "ActionEndDate": "4712-12-31",
            },
            {
                "ActionCode": "DEPARTMENT_CHANGE",
                "ActionName": "Employee Department Change",
                "ActionDescription": "Change of employee department",
                "ActionTypeCode": "EMPL_TRANSFER",
                "ActionStartDate": "2026-11-16",
                "ActionEndDate": "4712-12-31",
            },
            {
                "ActionCode": "SALARY_REVISION",
                "ActionName": "Employee Salary Revision",
                "ActionDescription": "Revision of employee salary",
                "ActionTypeCode": "EMPL_CHANGE",
                "ActionStartDate": "2026-11-16",
                "ActionEndDate": "4712-12-31",
            },
        ])

    def test_builds_distinct_actions_and_leaves_unprovided_ids_blank(self):
        output_objects = build_actions_output_objects(
            self.dataframe,
            "GRISMO",
        )
        records = output_objects[0]["records"]

        self.assertEqual(len(records), 2)
        self.assertEqual(
            records[0],
            {
                "ActionId": "",
                "ActionCode": "SALARY_REVISION",
                "ActionName": "Employee Salary Revision",
                "Description": "Revision of employee salary",
                "ActionTypeId": "",
                "ActionTypeCode": "EMPL_CHANGE",
                "StartDate": "2026-11-16",
                "EndDate": "4712-12-31",
                "GUID": "",
                "SourceSystemId": "ACT_SALARY_REVISION_001",
                "SourceSystemOwner": "GRISMO",
            },
        )
        self.assertEqual(
            output_objects[0]["metadata_fields"][0],
            "FLEX:PER_ACT_LEG_DDF",
        )

    def test_rejects_missing_columns_and_incomplete_action_rows(self):
        with self.assertRaisesRegex(ValueError, "requires Excel columns"):
            build_actions_output_objects(
                pd.DataFrame({"ActionCode": ["HIRE"]}),
                "GRISMO",
            )

        incomplete = self.dataframe.iloc[:1].copy()
        incomplete.loc[0, "ActionTypeCode"] = None
        with self.assertRaisesRegex(ValueError, "Excel row 2"):
            build_actions_output_objects(incomplete, "GRISMO")

    def test_rejects_conflicting_definitions_for_same_action_code(self):
        conflicting = self.dataframe.iloc[:2].copy()
        conflicting.loc[1, "ActionCode"] = "SALARY_REVISION"

        with self.assertRaisesRegex(ValueError, "conflicting Actions data"):
            build_actions_output_objects(conflicting, "GRISMO")

    def test_registers_template_and_maps_workbook_aliases(self):
        with patch("frontend.service.get_templates", return_value=[]):
            templates = get_available_templates()

        template = next(
            item for item in templates if item["template_name"] == "Actions"
        )
        self.assertEqual(template["output_file_name"], "Actions.dat")
        self.assertEqual(template["category"], "Organization Data")

        with patch("frontend.service.read_excel", return_value=self.dataframe):
            mappings = analyze_field_mappings(
                "input.xlsx",
                "Actions",
            )["mappings"]

        description_mapping = next(
            item for item in mappings if item["field_name"] == "Description"
        )
        self.assertEqual(
            description_mapping["excel_column"],
            "ActionDescription",
        )
        self.assertEqual(description_mapping["status"], "MAPPED")

    def test_pipeline_writes_actions_dat(self):
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
                    template_name="Actions",
                    output_dir=str(output_dir),
                    reports_dir=str(reports_dir),
                )

            content = Path(result["output_file"]).read_text(
                encoding="utf-8",
            )

        self.assertEqual(result["generated_rows"], 2)
        self.assertIn(
            "METADATA|Actions|FLEX:PER_ACT_LEG_DDF|ActionId|ActionCode|"
            "ActionName|Description|ActionTypeId|ActionTypeCode|StartDate|"
            "EndDate|GUID|SourceSystemId|SourceSystemOwner",
            content,
        )
        self.assertIn(
            "MERGE|Actions|||SALARY_REVISION|Employee Salary Revision|"
            "Revision of employee salary||EMPL_CHANGE|2026/11/16|"
            "4712/12/31||ACT_SALARY_REVISION_001|GRISMO",
            content,
        )


if __name__ == "__main__":
    unittest.main()
