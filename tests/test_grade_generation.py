from pathlib import Path
import contextlib
import io
import sys
import tempfile
import unittest
from unittest.mock import patch

import pandas as pd

import main
from frontend.service import (
    analyze_field_mappings,
    get_available_templates,
    run_hdl_generation_pipeline,
)
from src.grade_generator import build_grade_output_objects


class GradeGenerationTests(unittest.TestCase):
    def setUp(self):
        self.dataframe = pd.DataFrame([
            {
                "GradeEffectiveStartDate": "1951-01-01",
                "EffectiveEndDate": None,
                "SetCode": None,
                "GradeCode": "Tech_Grade",
                "GradeName": "TG1",
                "ActiveStatus": None,
            },
            {
                "GradeEffectiveStartDate": "1951-01-01",
                "EffectiveEndDate": "4712-12-31",
                "SetCode": "COMMON",
                "GradeCode": "Tech_Grade2",
                "GradeName": "TG2",
                "ActiveStatus": "A",
            },
            {
                "GradeEffectiveStartDate": "1951-01-01",
                "EffectiveEndDate": None,
                "SetCode": None,
                "GradeCode": "Tech_Grade",
                "GradeName": "TG1",
                "ActiveStatus": None,
            },
        ])

    def test_builds_distinct_grades_with_sample_defaults(self):
        output_objects = build_grade_output_objects(
            self.dataframe,
            "GRISMO",
        )

        self.assertEqual(len(output_objects[0]["records"]), 2)
        self.assertEqual(output_objects[0]["records"][0], {
            "SourceSystemOwner": "GRISMO",
            "SourceSystemId": "GRD_TECH_GRADE",
            "EffectiveStartDate": "1951/01/01",
            "EffectiveEndDate": "4712/12/31",
            "SetCode": "COMMON",
            "GradeCode": "Tech_Grade",
            "GradeName": "TG1",
            "ActiveStatus": "A",
        })

    def test_allows_blank_source_system_owner(self):
        output_objects = build_grade_output_objects(self.dataframe, "")
        self.assertEqual(
            output_objects[0]["records"][0]["SourceSystemOwner"],
            "",
        )

    def test_rejects_missing_columns_and_incomplete_rows(self):
        with self.assertRaisesRegex(ValueError, "requires Excel columns"):
            build_grade_output_objects(
                pd.DataFrame({"GradeCode": ["Tech_Grade"]}),
                "GRISMO",
            )

        incomplete = self.dataframe.iloc[:1].copy()
        incomplete.loc[0, "GradeName"] = None
        with self.assertRaisesRegex(ValueError, "Excel row 2"):
            build_grade_output_objects(incomplete, "GRISMO")

    def test_registers_template_and_maps_grade_specific_start_date(self):
        with patch("frontend.service.get_templates", return_value=[]):
            templates = get_available_templates()
        template = next(
            item for item in templates if item["template_name"] == "Grade"
        )
        self.assertEqual(template["output_file_name"], "Grade.dat")

        with patch("frontend.service.read_excel", return_value=self.dataframe):
            mappings = analyze_field_mappings(
                "input.xlsx",
                "Grade",
            )["mappings"]

        start_mapping = next(
            item for item in mappings
            if item["field_name"] == "EffectiveStartDate"
        )
        self.assertEqual(
            start_mapping["excel_column"],
            "GradeEffectiveStartDate",
        )
        self.assertEqual(start_mapping["status"], "MAPPED")

    def test_pipeline_writes_grade_dat(self):
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
                    template_name="Grade",
                    output_dir=str(output_dir),
                    reports_dir=str(reports_dir),
                )

            content = Path(result["output_file"]).read_text(
                encoding="utf-8",
            )

        self.assertEqual(result["generated_rows"], 2)
        self.assertIn(
            "METADATA|Grade|SourceSystemOwner|SourceSystemId|"
            "EffectiveStartDate|EffectiveEndDate|SetCode|GradeCode|"
            "GradeName|ActiveStatus",
            content,
        )
        self.assertIn(
            "MERGE|Grade|GRISMO|GRD_TECH_GRADE|1951/01/01|"
            "4712/12/31|COMMON|Tech_Grade|TG1|A",
            content,
        )

    def test_main_generates_grade_without_worker_template_mapping(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output_dir = Path(temp_dir) / "output"
            args = [
                "main.py",
                "--input",
                "input.xlsx",
                "--template",
                "Grade",
                "--output-dir",
                str(output_dir),
            ]
            with (
                patch.object(sys, "argv", args),
                patch("main.read_excel", return_value=self.dataframe),
                patch(
                    "main.get_active_clients",
                    return_value=[{"source_system_owner": "GRISMO"}],
                ),
                contextlib.redirect_stdout(io.StringIO()),
            ):
                main.main()

            output_file = output_dir / "Grade.dat"
            self.assertTrue(output_file.is_file())
            self.assertEqual(
                sum(
                    line.startswith("MERGE|")
                    for line in output_file.read_text(
                        encoding="utf-8",
                    ).splitlines()
                ),
                2,
            )

    def test_main_auto_detects_grade_when_template_is_omitted(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output_dir = Path(temp_dir) / "output"
            args = [
                "main.py",
                "--input",
                "input.xlsx",
                "--output-dir",
                str(output_dir),
            ]
            with (
                patch.object(sys, "argv", args),
                patch("main.read_excel", return_value=self.dataframe),
                patch(
                    "main.get_active_clients",
                    return_value=[{"source_system_owner": "GRISMO"}],
                ),
                patch("main.load_template_configuration") as load_template,
                contextlib.redirect_stdout(io.StringIO()),
            ):
                main.main()

            output_file = output_dir / "Grade.dat"
            self.assertTrue(output_file.is_file())
            load_template.assert_not_called()
            self.assertEqual(
                sum(
                    line.startswith("MERGE|")
                    for line in output_file.read_text(
                        encoding="utf-8",
                    ).splitlines()
                ),
                2,
            )

    def test_default_worker_template_error_suggests_grade_option(self):
        args = [
            "main.py",
            "--input",
            "input.xlsx",
            "--template",
            "Worker",
        ]
        with (
            patch.object(sys, "argv", args),
            patch(
                "main.load_template_configuration",
                return_value=(
                    {
                        "template_name": "Worker",
                        "source_system_owner": "GRISMO",
                    },
                    [],
                ),
            ),
            patch("main.read_excel", return_value=self.dataframe),
            patch("main.load_field_mappings", return_value={}),
            patch("main.map_excel_dataframe", return_value=([{}], [])),
            contextlib.redirect_stdout(io.StringIO()),
        ):
            with self.assertRaisesRegex(SystemExit, "--template Grade"):
                main.main()


if __name__ == "__main__":
    unittest.main()
