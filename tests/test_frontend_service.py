from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import pandas as pd

from frontend.service import (
    _person_disability_validation_error,
    delete_input_file,
    inspect_excel_data,
    run_hdl_generation_pipeline,
)


class ExcelPreviewTests(unittest.TestCase):
    def test_serializes_missing_datetime_cells_as_blank(self):
        dataframe = pd.DataFrame({
            "StartDate": [
                pd.Timestamp("2026-01-01"),
                pd.NaT,
            ],
            "Description": ["Available", None],
        })

        with patch("frontend.service.read_excel", return_value=dataframe):
            result = inspect_excel_data("input.xlsx", page_size=10)

        self.assertEqual(result["rows"], [
            {
                "_row_number": 2,
                "StartDate": "2026-01-01",
                "Description": "Available",
            },
            {
                "_row_number": 3,
                "StartDate": "",
                "Description": "",
            },
        ])


class InputWorkbookDeletionTests(unittest.TestCase):
    def test_deletes_workbook_from_input_directory(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_dir = Path(temp_dir) / "input"
            input_dir.mkdir()
            workbook = input_dir / "unwanted.xlsx"
            workbook.write_bytes(b"workbook")

            result = delete_input_file(workbook.name, str(input_dir))

            self.assertEqual(result, {
                "success": True,
                "filename": "unwanted.xlsx",
            })
            self.assertFalse(workbook.exists())

    def test_rejects_paths_and_non_workbook_files(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_dir = Path(temp_dir) / "input"
            input_dir.mkdir()
            outside_file = Path(temp_dir) / "outside.xlsx"
            outside_file.write_bytes(b"keep")

            with self.assertRaisesRegex(ValueError, "not a path"):
                delete_input_file("../outside.xlsx", str(input_dir))
            with self.assertRaisesRegex(ValueError, "Only Excel workbooks"):
                delete_input_file("output.dat", str(input_dir))

            self.assertTrue(outside_file.exists())

    def test_reports_missing_workbook(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_dir = Path(temp_dir) / "input"
            input_dir.mkdir()

            with self.assertRaises(FileNotFoundError):
                delete_input_file("missing.xls", str(input_dir))


class PersonDisabilityValidationErrorTests(unittest.TestCase):
    def test_converts_row_data_errors_to_validation_issues(self):
        error = _person_disability_validation_error(ValueError(
            "Excel row 13: disability details 'Mobility impairment' "
            "did not match an enabled person_disability_lookups value."
        ))

        self.assertEqual(error, {
            "ExcelRow": 13,
            "Object": "PersonDisability",
            "Field": "Disability",
            "Error": (
                "disability details 'Mobility impairment' did not match an "
                "enabled person_disability_lookups value."
            ),
        })

    def test_does_not_convert_non_row_validation_errors(self):
        self.assertIsNone(_person_disability_validation_error(ValueError(
            "Unable to load disability lookup configuration."
        )))


class ContactGenerationPipelineTests(unittest.TestCase):
    def test_contact_template_uses_generated_family_contact_records(self):
        dataframe = pd.DataFrame([{
            "Person Number": "SAEMP1001",
            "Start Date": "2020-01-01",
            "Father Name": "Rajiv Sharma",
            "Father Email": "father@example.test",
            "Home Address": "1 Main Street",
            "Nationality": "Saudi Arabian",
        }])
        template = {
            "template_name": "Contact",
            "source_system_owner": "GRISMO",
        }

        with tempfile.TemporaryDirectory() as temp_dir:
            output_dir = Path(temp_dir) / "output"
            reports_dir = Path(temp_dir) / "reports"
            with (
                patch(
                    "frontend.service.load_template_configuration",
                    return_value=(template, []),
                ),
                patch("frontend.service.read_excel", return_value=dataframe),
            ):
                result = run_hdl_generation_pipeline(
                    "input.xlsx",
                    template_name="Contact",
                    output_dir=str(output_dir),
                    reports_dir=str(reports_dir),
                )

            content = Path(result["output_file"]).read_text(
                encoding="utf-8",
            )
            self.assertIn(
                "MERGE|Contact||2020/01/01|||||2020/01/01",
                content,
            )
            self.assertIn("PER_FAT_SAEMP1001", content)
            self.assertIn("Rajiv", content)
            self.assertEqual(result["generated_rows"], 6)
            self.assertEqual(result["validation_errors_count"], 0)
            self.assertEqual(result["validation_errors"], [])
            self.assertEqual(
                Path(result["validation_report_path"]).read_text(
                    encoding="utf-8",
                ).strip(),
                "ExcelRow,Object,Field,Error",
            )

    def test_contact_template_fails_when_no_family_contacts_are_generated(self):
        dataframe = pd.DataFrame([{
            "Person Number": "SAEMP1001",
            "Start Date": "2020-01-01",
        }])
        template = {
            "template_name": "Contact",
            "source_system_owner": "GRISMO",
        }

        with tempfile.TemporaryDirectory() as temp_dir:
            with (
                patch(
                    "frontend.service.load_template_configuration",
                    return_value=(template, []),
                ),
                patch("frontend.service.read_excel", return_value=dataframe),
            ):
                with self.assertRaisesRegex(
                    ValueError,
                    "No Contact records were generated",
                ):
                    run_hdl_generation_pipeline(
                        "input.xlsx",
                        template_name="Contact",
                        output_dir=str(Path(temp_dir) / "output"),
                        reports_dir=str(Path(temp_dir) / "reports"),
                    )


if __name__ == "__main__":
    unittest.main()
