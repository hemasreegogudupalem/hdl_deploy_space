from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import pandas as pd

from frontend.service import run_hdl_generation_pipeline


class ResourceExceptionGenerationTests(unittest.TestCase):
    def test_generates_mapped_records_with_datetime_and_source_ids(self):
        dataframe = pd.DataFrame([
            {
                "ExceptionName": "Medical Appointment",
                "StartDateTime": "2026/11/20 09:00:00",
                "EndDateTime": "2026/11/20 12:00:00",
                "AvailabilityCode": "OFF_PERIOD",
            },
            {
                "ExceptionName": "Training Session",
                "StartDateTime": "2026/11/22 10:00:00",
                "EndDateTime": "2026/11/22 16:00:00",
                "AvailabilityCode": "WORK_PERIOD",
            },
        ])
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
                    template_name="ResourceException",
                    output_dir=str(output_dir),
                    reports_dir=str(reports_dir),
                )

            content = Path(result["output_file"]).read_text(encoding="utf-8")

        field_names = [
            "ExceptionId",
            "ExceptionName",
            "StartDateTime",
            "EndDateTime",
            "AvailabilityCode",
            "SourceSystemOwner",
            "SourceSystemId",
        ]
        self.assertIn(
            "METADATA|ResourceException|" + "|".join(field_names),
            content,
        )
        self.assertIn(
            "MERGE|ResourceException||Medical Appointment|"
            "2026/11/20 09:00:00|2026/11/20 12:00:00|OFF_PERIOD|"
            "GRISMO|RES_EXC_001",
            content,
        )
        self.assertIn(
            "MERGE|ResourceException||Training Session|"
            "2026/11/22 10:00:00|2026/11/22 16:00:00|WORK_PERIOD|"
            "GRISMO|RES_EXC_002",
            content,
        )
        self.assertEqual(result["generated_rows"], 2)


if __name__ == "__main__":
    unittest.main()
