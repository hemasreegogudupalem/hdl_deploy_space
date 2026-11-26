from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import pandas as pd

from frontend.service import run_hdl_generation_pipeline
from src.person_type_generator import build_person_type_output_objects


class PersonTypeGenerationTests(unittest.TestCase):
    def test_builds_one_record_for_each_distinct_definition(self):
        dataframe = pd.DataFrame([
            {
                "SystemPersonType": "EMP",
                "UserPersonType": "Employee",
                "ActiveFlag": "Y",
                "DefaultFlag": "Y",
            },
            {
                "SystemPersonType": "EMP",
                "UserPersonType": "Permanent Employee",
                "ActiveFlag": "Y",
                "DefaultFlag": "N",
            },
            {
                "SystemPersonType": "EMP",
                "UserPersonType": "Employee",
                "ActiveFlag": "Y",
                "DefaultFlag": "Y",
            },
            {
                "SystemPersonType": "CWK",
                "UserPersonType": "Contingent Worker",
                "ActiveFlag": "Y",
                "DefaultFlag": "Y",
            },
        ])

        output_objects = build_person_type_output_objects(
            dataframe,
            "GRISMO",
        )

        self.assertEqual(output_objects[0]["records"], [
            {
                "PersonTypeId": "",
                "SystemPersonType": "EMP",
                "UserPersonType": "Employee",
                "ActiveFlag": "Y",
                "DefaultFlag": "Y",
                "SourceSystemOwner": "GRISMO",
                "SourceSystemId": "NEW_EMP_002",
            },
            {
                "PersonTypeId": "",
                "SystemPersonType": "EMP",
                "UserPersonType": "Permanent Employee",
                "ActiveFlag": "Y",
                "DefaultFlag": "N",
                "SourceSystemOwner": "GRISMO",
                "SourceSystemId": "NEW_EMP_001",
            },
            {
                "PersonTypeId": "",
                "SystemPersonType": "CWK",
                "UserPersonType": "Contingent Worker",
                "ActiveFlag": "Y",
                "DefaultFlag": "Y",
                "SourceSystemOwner": "GRISMO",
                "SourceSystemId": "NEW_CWK_001",
            },
        ])

    def test_reports_missing_person_type_columns(self):
        with self.assertRaisesRegex(ValueError, "requires Excel columns"):
            build_person_type_output_objects(
                pd.DataFrame({"Worker Type": ["Seasonal Worker"]}),
                "GRISMO",
            )

    def test_person_type_pipeline_writes_hdl_file(self):
        dataframe = pd.DataFrame({
            "SystemPersonType": ["EMP", "EMP"],
            "UserPersonType": ["Employee", "Permanent Employee"],
            "ActiveFlag": ["Y", "Y"],
            "DefaultFlag": ["Y", "N"],
        })

        with tempfile.TemporaryDirectory() as temp_dir:
            output_dir = Path(temp_dir) / "output"
            reports_dir = Path(temp_dir) / "reports"
            with (
                patch(
                    "frontend.service.get_active_clients",
                    return_value=[{"source_system_owner": "GRISMO"}],
                ),
                patch("frontend.service.read_excel", return_value=dataframe),
            ):
                result = run_hdl_generation_pipeline(
                    "input.xlsx",
                    template_name="PersonType",
                    output_dir=str(output_dir),
                    reports_dir=str(reports_dir),
                )

            content = Path(result["output_file"]).read_text(encoding="utf-8")

        self.assertEqual(result["generated_rows"], 2)
        self.assertIn(
            "METADATA|PersonType|PersonTypeId|SystemPersonType|UserPersonType|"
            "ActiveFlag|DefaultFlag|SourceSystemOwner|SourceSystemId",
            content,
        )
        self.assertIn(
            "MERGE|PersonType||EMP|Permanent Employee|Y|N|GRISMO|NEW_EMP_001",
            content,
        )


if __name__ == "__main__":
    unittest.main()
