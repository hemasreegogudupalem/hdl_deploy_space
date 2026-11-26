import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

import pandas as pd

from frontend import service
from src.job_family_generator import (
    build_job_family_records,
    generate_job_family,
)


class JobFamilyGeneratorTests(unittest.TestCase):
    def setUp(self):
        self.dataframe = pd.DataFrame([
            {
                "JobFamilyCode": "CLOUD_FAMILY",
                "JobFamilyName": "Cloud Computing",
                "JobFamilyEffectiveStartDate": "2026/01/01",
                "JobSubFamily": "MIS_FAMILY",
            },
            {
                "JobFamilyCode": "CLOUD_FAMILY",
                "JobFamilyName": "Cloud Computing",
                "JobFamilyEffectiveStartDate": "2026/01/01",
                "JobSubFamily": "MIS_FAMILY",
            },
            {
                "JobFamilyCode": "DATA_FAMILY",
                "JobFamilyName": "Data Analytics",
                "JobFamilyEffectiveStartDate": "2026/01/01",
                "JobSubFamily": "ANALYTICS_FAMILY",
            },
        ])

    def test_builds_distinct_family_and_association_records(self):
        families, associations = build_job_family_records(self.dataframe)

        self.assertEqual(len(families), 2)
        self.assertEqual(len(associations), 2)
        self.assertEqual(families[0]["EffectiveStartDate"], "2026/01/01")
        self.assertEqual(families[0]["EffectiveEndDate"], "4712/12/31")
        self.assertEqual(associations[0]["JobSubFamily"], "MIS_FAMILY")

    def test_generation_pipeline_writes_jobfamily_output(self):
        with tempfile.TemporaryDirectory() as output_dir:
            with mock.patch.object(
                service,
                "read_excel",
                return_value=self.dataframe,
            ):
                result = service.run_hdl_generation_pipeline(
                    "input/saudi_employee_data_job.xlsx",
                    template_name="JobFamily",
                    output_dir=output_dir,
                )

            output_file = Path(output_dir) / "JobFamily.dat"
            self.assertEqual(result["output_filename"], "JobFamily.dat")
            self.assertEqual(
                result["output_download_url"],
                "/api/reports/download/output/JobFamily.dat",
            )
            self.assertEqual(result["template_name"], "JobFamily")
            self.assertEqual(result["generated_rows"], 4)
            self.assertTrue(output_file.exists())
            output = output_file.read_text(encoding="utf-8")
            self.assertIn("MERGE|JobFamily|DATA_FAMILY|Data Analytics", output)
            self.assertIn(
                "MERGE|JobSubFamilyAssociation|ANALYTICS_FAMILY|"
                "Data Analytics|DATA_FAMILY",
                output,
            )

    def test_jobfamily_template_available_without_database(self):
        with mock.patch.object(
            service,
            "get_templates",
            side_effect=RuntimeError("database unavailable"),
        ):
            templates = service.get_available_templates()

        template_names = {template["template_name"] for template in templates}
        self.assertIn("JobFamily", template_names)
        self.assertIn("ResourceException", template_names)
        self.assertIn("PersonType", template_names)
        self.assertEqual(
            service.detect_generation_defaults("unused.xlsx", "JobFamily"),
            {
                "legislation_code": "",
                "legal_employer_name": "",
                "source_system_owner": "",
            },
        )

    def test_builtin_template_details_include_resource_exception_and_person_type(self):
        resource_exception = service.get_template_details("ResourceException")
        person_type = service.get_template_details("PersonType")

        self.assertEqual(
            resource_exception["output_file_name"],
            "ResourceException.dat",
        )
        self.assertEqual(
            resource_exception["objects"][0]["object_name"],
            "ResourceException",
        )
        self.assertEqual(
            person_type["output_file_name"],
            "PersonType.dat",
        )
        self.assertEqual(
            person_type["objects"][0]["object_name"],
            "PersonType",
        )
        self.assertIsNone(person_type["error"])

    def test_builtin_template_mapping_previews_use_excel_columns(self):
        person_type_frame = pd.DataFrame([{
            "SystemPersonType": "EMP",
            "UserPersonType": "Permanent Employee",
            "ActiveFlag": "Y",
            "DefaultFlag": "N",
        }])
        exception_frame = pd.DataFrame([{
            "ExceptionName": "Medical Appointment",
            "StartDateTime": "2026/11/20 09:00:00",
            "EndDateTime": "2026/11/20 12:00:00",
            "AvailabilityCode": "OFF_PERIOD",
        }])

        with mock.patch.object(
            service,
            "read_excel",
            side_effect=[person_type_frame, exception_frame],
        ):
            person_type = service.analyze_field_mappings(
                "person-types.xlsx",
                "PersonType",
            )
            exception = service.analyze_field_mappings(
                "exceptions.xlsx",
                "ResourceException",
            )

        self.assertEqual(person_type["mapped_count"], 7)
        self.assertEqual(exception["mapped_count"], 7)
        person_type_mapping = {
            item["field_name"]: item
            for item in person_type["mappings"]
        }
        self.assertEqual(
            person_type_mapping["UserPersonType"]["excel_column"],
            "UserPersonType",
        )
        self.assertEqual(
            person_type_mapping["SourceSystemId"]["mapping_type"],
            "GENERATED",
        )

    def test_builtin_template_defaults_use_active_client_owner(self):
        with mock.patch.object(
            service,
            "get_active_clients",
            return_value=[{"source_system_owner": "GRISMO"}],
        ):
            defaults = service.detect_generation_defaults(
                "unused.xlsx",
                "PersonType",
            )

        self.assertEqual(defaults["source_system_owner"], "GRISMO")

    def test_batch_validation_applies_custom_source_column_mappings(self):
        dataframe = pd.DataFrame([{
            "Family Code": "CLOUD_FAMILY",
            "Family Name": "Cloud Computing",
            "Start Date": "2026/01/01",
            "Subfamily": "MIS_FAMILY",
        }])
        selection = {
            "template_name": "JobFamily",
            "column_mappings": {
                "JobFamilyCode": "Family Code",
                "JobFamilyName": "Family Name",
                "EffectiveStartDate": "Start Date",
                "JobSubFamily": "Subfamily",
            },
        }

        with mock.patch.object(
            service,
            "read_excel",
            return_value=dataframe,
        ):
            result = service.validate_batch_generation(
                "custom-job-families.xlsx",
                [selection],
            )

        self.assertTrue(result["can_generate"])
        self.assertEqual(result["templates"][0]["generated_rows"], 2)
        self.assertTrue(result["templates"][0]["ready"])

    def test_batch_generation_creates_downloadable_zip_with_dat(self):
        dataframe = pd.DataFrame([{
            "JobFamilyCode": "CLOUD_FAMILY",
            "JobFamilyName": "Cloud Computing",
            "EffectiveStartDate": "2026/01/01",
            "JobSubFamily": "MIS_FAMILY",
        }])
        selection = {
            "template_name": "JobFamily",
            "column_mappings": {
                "JobFamilyCode": "JobFamilyCode",
                "JobFamilyName": "JobFamilyName",
                "EffectiveStartDate": "EffectiveStartDate",
                "JobSubFamily": "JobSubFamily",
            },
        }

        with tempfile.TemporaryDirectory() as output_dir:
            with mock.patch.object(
                service,
                "read_excel",
                return_value=dataframe,
            ):
                result = service.generate_batch_hdl(
                    "job-families.xlsx",
                    [selection],
                    output_dir=output_dir,
                )

            zip_path = Path(output_dir) / result["zip_filename"]
            self.assertTrue(zip_path.is_file())
            self.assertEqual(
                result["zip_download_url"],
                f"/api/reports/download/output/{result['zip_filename']}",
            )
            self.assertIn("JobFamily.dat", result["output_filenames"])
            listed_outputs = service.list_reports_and_outputs(
                output_dir=output_dir,
            )["outputs"]
            self.assertEqual(
                {item["filename"] for item in listed_outputs},
                {"JobFamily.dat", result["zip_filename"]},
            )
            with zipfile.ZipFile(zip_path) as archive:
                self.assertEqual(archive.namelist(), ["JobFamily.dat"])
                self.assertIn(
                    b"MERGE|JobFamily|CLOUD_FAMILY|Cloud Computing",
                    archive.read("JobFamily.dat"),
                )

    def test_batch_deduplication_ignores_only_source_system_id(self):
        dat_content = (
            "METADATA|Example|SourceSystemOwner|SourceSystemId|Name\n"
            "MERGE|Example|OWNER|ID_1|Same Name\n"
            "MERGE|Example|OWNER|ID_2|Same Name\n"
            "MERGE|Example|OTHER_OWNER|ID_3|Same Name\n"
        )
        with tempfile.TemporaryDirectory() as temporary_directory:
            output_file = Path(temporary_directory) / "Example.dat"
            output_file.write_text(dat_content, encoding="utf-8")

            removed_count = service._deduplicate_hdl_dat_file(output_file)

            self.assertEqual(removed_count, 1)
            self.assertEqual(
                output_file.read_text(encoding="utf-8").splitlines(),
                [
                    "METADATA|Example|SourceSystemOwner|SourceSystemId|Name",
                    "MERGE|Example|OWNER|ID_1|Same Name",
                    "MERGE|Example|OTHER_OWNER|ID_3|Same Name",
                ],
            )

    def test_batch_load_order_handles_compact_template_names(self):
        self.assertLess(
            service._batch_load_order("JobFamily"),
            service._batch_load_order("Worker"),
        )
        self.assertLess(
            service._batch_load_order("Worker"),
            service._batch_load_order("ResourceException"),
        )
        self.assertLess(
            service._batch_load_order("CalendarEvent"),
            service._batch_load_order("Worker"),
        )
        self.assertLess(
            service._batch_load_order("Action"),
            service._batch_load_order("Worker"),
        )

    def test_batch_schema_blocks_required_internal_id_mapping(self):
        template = {"template_name": "Example"}
        objects = [{
            "object_name": "ExampleObject",
            "fields": [
                {"field_name": "ExampleID", "required_flag": "Y"},
                {
                    "field_name": "RelatedObjectId(SourceSystemId)",
                    "required_flag": "N",
                },
            ],
            "mappings": [{
                "field_name": "ExampleID",
                "mapping_type": "COLUMN",
                "excel_column": "Internal ID",
            }],
        }]

        with mock.patch.object(
            service,
            "load_template_configuration",
            return_value=(template, objects),
        ):
            schema = service.get_batch_template_schema("Example")

        internal_id = schema["source_fields"][0]
        external_reference = schema["source_fields"][1]
        self.assertTrue(internal_id["internal_id"])
        self.assertFalse(internal_id["required"])
        self.assertTrue(internal_id["target_required"])
        self.assertFalse(external_reference["internal_id"])
        self.assertFalse(external_reference["required"])

        with self.assertRaisesRegex(ValueError, "Internal ID field"):
            service._apply_batch_mapping_overrides(
                objects,
                {"ExampleObject.ExampleID": "Internal ID"},
            )


if __name__ == "__main__":
    unittest.main()
