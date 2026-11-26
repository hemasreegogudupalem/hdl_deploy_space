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
from src.action_reason_generator import build_action_reason_output_objects


class ActionReasonGenerationTests(unittest.TestCase):
    def setUp(self):
        self.dataframe = pd.DataFrame([
            {
                "ActionReasonCode": "DIFF_PROJECT_CHANGE",
                "ActionReason": "Project Change",
                "ActionReasonStartDate": "2026-01-01",
                "ActionReasonEndDate": "4712-12-31",
            },
            {
                "ActionReasonCode": "PERFORMANCE_INCREASE",
                "ActionReason": "Performance Increase",
                "ActionReasonStartDate": "2026-01-01",
                "ActionReasonEndDate": "4712-12-31",
            },
            {
                "ActionReasonCode": "DIFF_PROJECT_CHANGE",
                "ActionReason": "Project Change",
                "ActionReasonStartDate": "2026-01-01",
                "ActionReasonEndDate": "4712-12-31",
            },
        ])

    def test_builds_distinct_reasons_with_sample_source_ids(self):
        output_objects = build_action_reason_output_objects(
            self.dataframe,
            "GRISMO",
        )

        self.assertEqual(output_objects[0]["records"], [
            {
                "SourceSystemOwner": "GRISMO",
                "SourceSystemId": "GRISMO_AR_101",
                "ActionReasonId": "",
                "ActionReasonCode": "DIFF_PROJECT_CHANGE",
                "ActionReason": "Project Change",
                "StartDate": "2026-01-01",
                "EndDate": "4712-12-31",
            },
            {
                "SourceSystemOwner": "GRISMO",
                "SourceSystemId": "GRISMO_AR_102",
                "ActionReasonId": "",
                "ActionReasonCode": "PERFORMANCE_INCREASE",
                "ActionReason": "Performance Increase",
                "StartDate": "2026-01-01",
                "EndDate": "4712-12-31",
            },
        ])

    def test_rejects_missing_columns_and_conflicting_reasons(self):
        with self.assertRaisesRegex(ValueError, "requires Excel columns"):
            build_action_reason_output_objects(
                pd.DataFrame({"ActionReason": ["Project Change"]}),
                "GRISMO",
            )

        conflicting = self.dataframe.iloc[:2].copy()
        conflicting.loc[1, "ActionReasonCode"] = "DIFF_PROJECT_CHANGE"
        with self.assertRaisesRegex(
            ValueError,
            "conflicting ActionReasons data",
        ):
            build_action_reason_output_objects(conflicting, "GRISMO")

    def test_registers_template_and_maps_action_reason_dates(self):
        with patch("frontend.service.get_templates", return_value=[]):
            templates = get_available_templates()
        template = next(
            item for item in templates
            if item["template_name"] == "ActionReasons"
        )
        self.assertEqual(template["output_file_name"], "ActionReasons.dat")

        with patch("frontend.service.read_excel", return_value=self.dataframe):
            mappings = analyze_field_mappings(
                "input.xlsx",
                "ActionReasons",
            )["mappings"]
        start_mapping = next(
            item for item in mappings if item["field_name"] == "StartDate"
        )
        self.assertEqual(
            start_mapping["excel_column"],
            "ActionReasonStartDate",
        )
        self.assertEqual(start_mapping["status"], "MAPPED")

    def test_pipeline_writes_action_reasons_dat(self):
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
                    template_name="ActionReasons",
                    output_dir=str(output_dir),
                    reports_dir=str(reports_dir),
                )

            content = Path(result["output_file"]).read_text(
                encoding="utf-8",
            )

        self.assertEqual(result["generated_rows"], 2)
        self.assertIn(
            "METADATA|ActionReasons|SourceSystemOwner|SourceSystemId|"
            "ActionReasonId|ActionReasonCode|ActionReason|StartDate|EndDate",
            content,
        )
        self.assertIn(
            "MERGE|ActionReasons|GRISMO|GRISMO_AR_101||"
            "DIFF_PROJECT_CHANGE|Project Change|2026/01/01|4712/12/31",
            content,
        )


if __name__ == "__main__":
    unittest.main()
