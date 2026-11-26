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
from src.calendar_event_generator import build_calendar_event_output_objects


class CalendarEventGenerationTests(unittest.TestCase):
    def setUp(self):
        self.events = pd.DataFrame([
            {
                "Name": "New Year's Day 2026",
                "Description": "Public Holiday for New Year",
                "Category": "PH",
                "CoverageType": "G",
                "ShortCode": "NYD2026",
                "StartDateTime": "2026-01-01 00:00:00",
                "EndDateTime": "2026-01-01 23:59:59",
                "TreeCode": "ORA_HRX_SA_GEO_TREE",
                "TreeStructureCode": "PER_GEO_TREE_STRUCTURE",
                "TreeVersionName": "Geography Tree Version for Saudi Arabia",
                "HalfDayForElapsed": "N",
            },
            {
                "Name": "Republic Day",
                "Description": "Republic Day sample event",
                "Category": "PH",
                "CoverageType": "G",
                "ShortCode": "IN_REP2026",
                "StartDateTime": "2026-01-26 00:00:00",
                "EndDateTime": "2026-01-26 23:59:59",
                "TreeCode": None,
                "TreeStructureCode": None,
                "TreeVersionName": None,
                "HalfDayForElapsed": "N",
            },
        ])
        self.coverage = pd.DataFrame([
            {
                "ShortCode": "NYD2026",
                "CoverageScope": "I",
                "TerritoryCode": "SA",
                "TreeStructureCode": "PER_GEO_TREE_STRUCTURE",
                "TreeCode": "ORA_HRX_SA_GEO_TREE",
                "TreeVersionName": "Geography Tree Version for Saudi Arabia",
            },
            {
                "ShortCode": "IN_REP2026",
                "CoverageScope": "I",
                "TerritoryCode": "IN",
                "TreeStructureCode": None,
                "TreeCode": None,
                "TreeVersionName": None,
            },
        ])

    def test_builds_event_and_matching_coverage_objects(self):
        output_objects = build_calendar_event_output_objects(
            self.events,
            self.coverage,
        )

        self.assertEqual(
            [obj["object_name"] for obj in output_objects],
            ["CalendarEvent", "CalendarEventCoverage"],
        )
        self.assertEqual(len(output_objects[0]["records"]), 2)
        self.assertEqual(len(output_objects[1]["records"]), 2)
        self.assertEqual(
            output_objects[0]["records"][0]["ShortCode"],
            "NYD2026",
        )
        self.assertEqual(
            output_objects[1]["records"][0]["TreeCode"],
            "ORA_HRX_SA_GEO_TREE",
        )
        self.assertEqual(
            output_objects[0]["records"][1]["TreeCode"],
            "",
        )

    def test_rejects_missing_or_unmatched_coverage_records(self):
        with self.assertRaisesRegex(ValueError, "requires Excel columns"):
            build_calendar_event_output_objects(
                self.events.drop(columns=["Category"]),
                self.coverage,
            )

        with self.assertRaisesRegex(ValueError, "no CalendarEvent found"):
            build_calendar_event_output_objects(
                self.events,
                self.coverage.assign(ShortCode=["MISSING", "IN_REP2026"]),
            )

    def test_registers_template_and_maps_both_sheets(self):
        with patch("frontend.service.get_templates", return_value=[]):
            templates = get_available_templates()
        template = next(
            item for item in templates
            if item["template_name"] == "CalendarEvent"
        )
        self.assertEqual(template["output_file_name"], "CalendarEvent.dat")
        self.assertEqual(template["objects_count"], 2)

        details = get_template_details("CalendarEvent")
        self.assertEqual(
            [obj["object_name"] for obj in details["objects"]],
            ["CalendarEvent", "CalendarEventCoverage"],
        )

        with (
            patch("frontend.service.read_excel", return_value=self.events),
            patch(
                "frontend.service._read_excel_sheet",
                return_value=self.coverage,
            ),
        ):
            mappings = analyze_field_mappings(
                "input.xlsx",
                "CalendarEvent",
            )

        self.assertEqual(mappings["unmapped_count"], 0)
        self.assertTrue(any(
            item["object_name"] == "CalendarEventCoverage"
            and item["field_name"] == "CoverageScope"
            for item in mappings["mappings"]
        ))

    def test_pipeline_writes_both_hdl_objects(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output_dir = Path(temp_dir) / "output"
            reports_dir = Path(temp_dir) / "reports"
            with (
                patch("frontend.service.read_excel", return_value=self.events),
                patch(
                    "frontend.service._read_excel_sheet",
                    return_value=self.coverage,
                ),
            ):
                result = run_hdl_generation_pipeline(
                    "input.xlsx",
                    template_name="CalendarEvent",
                    output_dir=str(output_dir),
                    reports_dir=str(reports_dir),
                )

            content = Path(result["output_file"]).read_text(
                encoding="utf-8",
            )

        self.assertEqual(result["generated_rows"], 4)
        self.assertIn(
            "METADATA|CalendarEvent|Name|Description|Category|CoverageType|"
            "ShortCode|StartDateTime|EndDateTime|TreeCode|TreeStructureCode|"
            "TreeVersionName|HalfDayForElapsed",
            content,
        )
        self.assertIn(
            "MERGE|CalendarEvent|New Year's Day 2026|"
            "Public Holiday for New Year|PH|G|NYD2026|"
            "2026/01/01 00:00:00|2026/01/01 23:59:59|"
            "ORA_HRX_SA_GEO_TREE|PER_GEO_TREE_STRUCTURE|"
            "Geography Tree Version for Saudi Arabia|N",
            content,
        )
        self.assertIn(
            "MERGE|CalendarEventCoverage|NYD2026|I|SA|"
            "PER_GEO_TREE_STRUCTURE|ORA_HRX_SA_GEO_TREE|"
            "Geography Tree Version for Saudi Arabia",
            content,
        )


if __name__ == "__main__":
    unittest.main()
