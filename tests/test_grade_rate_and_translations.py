from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import pandas as pd

from frontend.service import (
    analyze_field_mappings,
    get_available_templates,
    get_template_details,
    run_hdl_generation_pipeline,
)
from src.action_reasons_translation_generator import (
    build_action_reasons_translation_output_objects,
)
from src.actions_translation_generator import (
    build_actions_translation_output_objects,
)
from src.grade_rate_generator import build_grade_rate_output_objects


class GradeRateAndTranslationGenerationTests(unittest.TestCase):
    def setUp(self):
        self.grade_rate_dataframe = pd.DataFrame([
            {
                "GradeRateStartDate": "2026-01-01",
                "GradeRateEndDate": "4712-12-31",
                "RateType": "Salary",
                "GradeRateName": "MIS Saudi Salary Rate",
                "GradeRateCurrencyCode": "SAR",
                "RateFrequency": "Monthly",
                "AnnualizationFactor": 12.0,
                "GradeRateLegislativeDataGroup": (
                    "SA Legislative Data Group"
                ),
                "GradeRate_ProgressionRateFlag": "N",
                "GradeRate_GradeCode": "E14",
                "GradeRate_MinimumAmount": 3000.0,
                "GradeRate_MaximumAmount": 5000.0,
                "GradeRate_MidValueAmount": 4000.0,
            },
            {
                "GradeRateStartDate": "2026-01-01",
                "GradeRateEndDate": "4712-12-31",
                "RateType": "Salary",
                "GradeRateName": "MIS Saudi Salary Rate",
                "GradeRateCurrencyCode": "SAR",
                "RateFrequency": "Monthly",
                "AnnualizationFactor": 12.0,
                "GradeRateLegislativeDataGroup": (
                    "SA Legislative Data Group"
                ),
                "GradeRate_ProgressionRateFlag": "N",
                "GradeRate_GradeCode": "E12",
                "GradeRate_MinimumAmount": 2500.0,
                "GradeRate_MaximumAmount": 4500.0,
                "GradeRate_MidValueAmount": 3500.0,
            },
        ])
        self.actions_translation_dataframe = pd.DataFrame([
            {
                "ActionsTranslation_ActionId": None,
                "ActionsTranslation_ActionCode": "SALARY_REVISION",
                "ActionsTranslation_ActionName": "Employee Salary Revision",
                "ActionsTranslation_Description": (
                    "Revision of employee salary"
                ),
                "ActionsTranslation_Language": "US",
                "ActionsTranslation_SourceSystemId": "ACT_SAL_REV_001",
                "ActionsTranslation_SourceSystemOwner": "GRISMO",
            },
            {
                "ActionsTranslation_ActionId": None,
                "ActionsTranslation_ActionCode": "SALARY_REVISION",
                "ActionsTranslation_ActionName": "مراجعة راتب الموظف",
                "ActionsTranslation_Description": "مراجعة راتب الموظف",
                "ActionsTranslation_Language": "AR",
                "ActionsTranslation_SourceSystemId": "ACT_SAL_REV_001",
                "ActionsTranslation_SourceSystemOwner": "GRISMO",
            },
        ])
        self.action_reasons_translation_dataframe = pd.DataFrame([
            {
                "ActionReasonsTranslation_ActionReasonId": None,
                "ActionReasonsTranslation_ActionReasonCode": (
                    "DIFF_PROJECT_CHANGE"
                ),
                "ActionReasonsTranslation_ActionReason": "Project Change",
                "ActionReasonsTranslation_Language": "US",
                "ActionReasonsTranslation_SourceSystemId": "GRISMO_AR_101",
                "ActionReasonsTranslation_SourceSystemOwner": "GRISMO",
            },
            {
                "ActionReasonsTranslation_ActionReasonId": None,
                "ActionReasonsTranslation_ActionReasonCode": (
                    "DIFF_PROJECT_CHANGE"
                ),
                "ActionReasonsTranslation_ActionReason": "تغيير المشروع",
                "ActionReasonsTranslation_Language": "AR",
                "ActionReasonsTranslation_SourceSystemId": "GRISMO_AR_101",
                "ActionReasonsTranslation_SourceSystemOwner": "GRISMO",
            },
        ])

    def test_builds_linked_grade_rate_and_value_rows(self):
        grade_rate, grade_rate_values = build_grade_rate_output_objects(
            self.grade_rate_dataframe,
            "GRISMO",
        )

        self.assertEqual(len(grade_rate["records"]), 1)
        self.assertEqual(len(grade_rate_values["records"]), 2)
        self.assertEqual(
            grade_rate["records"][0]["SourceSystemId"],
            "RATE_001",
        )
        self.assertEqual(
            grade_rate_values["records"][0]["RateId(SourceSystemId)"],
            "RATE_001",
        )
        self.assertEqual(
            grade_rate_values["records"][0]["MinimumAmount"],
            "3000",
        )
        self.assertEqual(
            grade_rate_values["records"][0]["SourceSystemId"],
            "RATE_VAL_001",
        )

    def test_builds_translation_records_and_preserves_languages(self):
        actions = build_actions_translation_output_objects(
            self.actions_translation_dataframe,
        )[0]["records"]
        reasons = build_action_reasons_translation_output_objects(
            self.action_reasons_translation_dataframe,
        )[0]["records"]

        self.assertEqual(len(actions), 2)
        self.assertEqual(actions[1]["ActionId"], "")
        self.assertEqual(actions[1]["Language"], "AR")
        self.assertEqual(actions[1]["ActionName"], "مراجعة راتب الموظف")
        self.assertEqual(len(reasons), 2)
        self.assertEqual(reasons[1]["ActionReasonId"], "")
        self.assertEqual(reasons[1]["ActionReason"], "تغيير المشروع")

    def test_registers_templates_and_grade_rate_object_mappings(self):
        with patch("frontend.service.get_templates", return_value=[]):
            templates = get_available_templates()
        names = {template["template_name"] for template in templates}
        self.assertTrue({
            "GradeRate",
            "ActionReasonsTranslation",
            "ActionsTranslation",
        }.issubset(names))
        details = get_template_details("GradeRate")
        self.assertEqual(
            [obj["object_name"] for obj in details["objects"]],
            ["GradeRate", "GradeRateValue"],
        )

        with patch(
            "frontend.service.read_excel",
            return_value=self.grade_rate_dataframe,
        ):
            mappings = analyze_field_mappings(
                "input.xlsx",
                "GradeRate",
            )["mappings"]
        grade_code_mapping = next(
            item for item in mappings
            if item["object_name"] == "GradeRateValue"
            and item["field_name"] == "GradeCode"
        )
        self.assertEqual(
            grade_code_mapping["excel_column"],
            "GradeRate_GradeCode",
        )
        self.assertEqual(grade_code_mapping["status"], "MAPPED")

    def test_pipelines_write_all_three_dat_files(self):
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
                    return_value=self.grade_rate_dataframe,
                ),
            ):
                rate_result = run_hdl_generation_pipeline(
                    "input.xlsx",
                    template_name="GradeRate",
                    output_dir=str(output_dir),
                    reports_dir=str(reports_dir),
                )
            with patch(
                "frontend.service.read_excel",
                return_value=self.actions_translation_dataframe,
            ):
                actions_result = run_hdl_generation_pipeline(
                    "input.xlsx",
                    template_name="ActionsTranslation",
                    output_dir=str(output_dir),
                    reports_dir=str(reports_dir),
                )
            with patch(
                "frontend.service.read_excel",
                return_value=self.action_reasons_translation_dataframe,
            ):
                reasons_result = run_hdl_generation_pipeline(
                    "input.xlsx",
                    template_name="ActionReasonsTranslation",
                    output_dir=str(output_dir),
                    reports_dir=str(reports_dir),
                )

            rate_content = Path(rate_result["output_file"]).read_text(
                encoding="utf-8",
            )
            actions_content = Path(actions_result["output_file"]).read_text(
                encoding="utf-8",
            )
            reasons_content = Path(reasons_result["output_file"]).read_text(
                encoding="utf-8",
            )

        self.assertEqual(rate_result["generated_rows"], 3)
        self.assertEqual(actions_result["generated_rows"], 2)
        self.assertEqual(reasons_result["generated_rows"], 2)
        self.assertIn(
            "MERGE|GradeRate||2026/01/01|4712/12/31||Salary|"
            "MIS Saudi Salary Rate|SAR|Monthly|12|A|"
            "SA Legislative Data Group||GRISMO|RATE_001|N||",
            rate_content,
        )
        self.assertIn(
            "MERGE|GradeRateValue||2026/01/01|4712/12/31||RATE_001|"
            "3000|5000|4000||MIS Saudi Salary Rate|COMMON|E14|"
            "SA Legislative Data Group|GRISMO|RATE_VAL_001",
            rate_content,
        )
        self.assertIn(
            "MERGE|ActionsTranslation||SALARY_REVISION|"
            "مراجعة راتب الموظف|مراجعة راتب الموظف|AR|"
            "ACT_SAL_REV_001|GRISMO",
            actions_content,
        )
        self.assertIn(
            "MERGE|ActionReasonsTranslation||DIFF_PROJECT_CHANGE|"
            "تغيير المشروع|AR|GRISMO_AR_101|GRISMO",
            reasons_content,
        )


if __name__ == "__main__":
    unittest.main()
