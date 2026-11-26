import tempfile
import unittest
from datetime import date
from pathlib import Path

from src.disability_generator import (
    PERSON_DISABILITY_FIELDS,
    build_person_disability_records,
    generate_person_disability,
    has_disability_data,
)


class PersonDisabilityGenerationTests(unittest.TestCase):
    LOOKUPS = [{
        "disability_code": "ORA_HRX_IN_DIS_J",
        "disability_name": "Hearing impairment - Deafness",
    }]
    LOOKUPS_WITH_DEFAULT = LOOKUPS + [{
        "disability_code": "ORA_HRX_OTHER_DISABILITY",
        "disability_name": "Other disability",
    }]

    def test_builds_record_using_database_code_and_excel_details(self):
        records = build_person_disability_records(
            [{
                "PersonNumber": 1005,
                "Disability": "Hearing impairment - Deafness",
                "StartDate": date(2025, 2, 12),
                "LegislationCode": "SA",
                "DisabilityDetails": "Partial hearing loss in left ear",
                "Degree": 25,
                "QuotaFTE": 1.0,
                "WorkRestriction": "Requires phone amplification headset",
                "Status": "A",
            }],
            self.LOOKUPS,
        )

        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["DisabilityCode"], "ORA_HRX_IN_DIS_J")
        self.assertEqual(records[0]["PersonNumber"], "1005")
        self.assertEqual(records[0]["EffectiveStartDate"], date(2025, 2, 12))
        self.assertEqual(records[0]["LegislationCode"], "SA")
        self.assertEqual(
            records[0]["Description"],
            "Partial hearing loss in left ear",
        )
        self.assertEqual(records[0]["Degree"], "25")
        self.assertEqual(records[0]["QuotaFTE"], "1")
        self.assertEqual(
            records[0]["WorkRestriction"],
            "Requires phone amplification headset",
        )
        self.assertEqual(records[0]["Status"], "A")

    def test_accepts_code_as_excel_disability_value(self):
        records = build_person_disability_records(
            [{"PersonNumber": "1005", "Disability": "ORA_HRX_IN_DIS_J"}],
            self.LOOKUPS,
        )

        self.assertEqual(records[0]["DisabilityCode"], "ORA_HRX_IN_DIS_J")

    def test_skips_blank_disability_rows(self):
        rows = [{"PersonNumber": "1005", "Disability": ""}]

        self.assertFalse(has_disability_data(rows))
        self.assertEqual(build_person_disability_records(rows, []), [])

    def test_skips_negative_no_disability_values(self):
        rows = [{"PersonNumber": "1005", "Disability": "No"}]

        self.assertFalse(has_disability_data(rows))
        self.assertEqual(build_person_disability_records(rows, self.LOOKUPS), [])

    def test_yes_flag_uses_disability_code(self):
        rows = [{"PersonNumber": "1005", "Disability": "Yes"}]
        rows[0]["DisabilityCode"] = "ORA_HRX_IN_DIS_J"

        self.assertTrue(has_disability_data(rows))
        records = build_person_disability_records(rows, self.LOOKUPS)
        self.assertEqual(records[0]["DisabilityCode"], "ORA_HRX_IN_DIS_J")

    def test_generates_file_for_yes_flag_with_disability_code(self):
        with tempfile.TemporaryDirectory() as output_dir:
            output_file = generate_person_disability(
                [{
                    "PersonNumber": "1005",
                    "Disability": "Yes",
                    "DisabilityCode": "ORA_HRX_IN_DIS_J",
                }],
                self.LOOKUPS,
                output_dir,
            )

            self.assertEqual(output_file.name, "PersonDisability.dat")
            self.assertIn(
                "MERGE|PersonDisability||ORA_HRX_IN_DIS_J|",
                output_file.read_text(encoding="utf-8"),
            )

    def test_generation_legislation_code_overrides_workbook_value(self):
        with tempfile.TemporaryDirectory() as output_dir:
            output_file = generate_person_disability(
                [{
                    "PersonNumber": "1005",
                    "Disability": "Hearing impairment - Deafness",
                    "LegislationCode": "FROM_WORKBOOK",
                }],
                self.LOOKUPS,
                output_dir,
                legislation_code="SAU",
            )

            merge_line = next(
                line
                for line in output_file.read_text(encoding="utf-8").splitlines()
                if line.startswith("MERGE|PersonDisability|")
            )
            self.assertEqual(merge_line.split("|")[6], "SAU")

    def test_generates_file_for_yes_flag_using_generic_database_code(self):
        with tempfile.TemporaryDirectory() as output_dir:
            output_file = generate_person_disability(
                [{
                    "PersonNumber": "1005",
                    "Disability": "Yes",
                }],
                self.LOOKUPS_WITH_DEFAULT,
                output_dir,
            )

            self.assertEqual(output_file.name, "PersonDisability.dat")
            self.assertIn(
                "MERGE|PersonDisability||ORA_HRX_OTHER_DISABILITY|",
                output_file.read_text(encoding="utf-8"),
            )

    def test_yes_flag_uses_generic_code_from_enabled_lookup(self):
        rows = [{
            "PersonNumber": "1005",
            "Disability": "Yes",
        }]

        self.assertTrue(has_disability_data(rows))
        records = build_person_disability_records(
            rows,
            self.LOOKUPS_WITH_DEFAULT,
        )
        self.assertEqual(
            records[0]["DisabilityCode"],
            "ORA_HRX_OTHER_DISABILITY",
        )

    def test_generic_code_is_taken_from_database_lookup_row(self):
        lookups = [{
            "disability_code": "DB_CONFIGURED_GENERIC_CODE",
            "disability_name": "Other disability",
        }]

        records = build_person_disability_records(
            [{"PersonNumber": "1005", "Disability": "Yes"}],
            lookups,
        )

        self.assertEqual(
            records[0]["DisabilityCode"],
            "DB_CONFIGURED_GENERIC_CODE",
        )

    def test_yes_flag_fails_when_generic_code_is_not_enabled(self):
        rows = [{"PersonNumber": "1005", "Disability": "Yes"}]

        with self.assertRaisesRegex(ValueError, "does not contain 'Other disability'"):
            build_person_disability_records(rows, self.LOOKUPS)

    def test_yes_flag_uses_details_when_they_match_a_lookup_value(self):
        rows = [{
            "PersonNumber": "1005",
            "Disability": "Yes",
            "DisabilityDetails": "Hearing impairment - Deafness",
        }]

        records = build_person_disability_records(rows, self.LOOKUPS)
        self.assertEqual(records[0]["DisabilityCode"], "ORA_HRX_IN_DIS_J")

    def test_yes_flag_matches_disability_details_contained_in_description(self):
        lookups = self.LOOKUPS + [
            {
                "disability_code": "SA_HEA_IMP",
                "disability_name": "Hearing impairment",
            },
            {
                "disability_code": "SA_PHY_DIS",
                "disability_name": "Physical disability",
            },
            {
                "disability_code": "SA_VIS_IMP",
                "disability_name": "Visual impairment",
            },
            {
                "disability_code": "SA_PSY_DIS",
                "disability_name": "Psychiatric disability",
            },
            {
                "disability_code": "SA_INT_DIS",
                "disability_name": "Intellectual disability",
            },
        ]
        rows = [
            {
                "PersonNumber": "1005",
                "Disability": "Yes",
                "Description": "Synthetic test data: hearing impairment.",
            },
            {
                "PersonNumber": "1006",
                "Disability": "Yes",
                "Description": "Synthetic test data: physical disability affecting mobility.",
            },
            {
                "PersonNumber": "1007",
                "Disability": "Yes",
                "Description": "Synthetic test data: visual impairment.",
            },
            {
                "PersonNumber": "1008",
                "Disability": "Yes",
                "Description": "Synthetic test data: psychiatric disability.",
            },
            {
                "PersonNumber": "1009",
                "Disability": "Yes",
                "Description": "Synthetic test data: intellectual disability.",
            },
        ]

        records = build_person_disability_records(rows, lookups)

        self.assertEqual(
            [record["DisabilityCode"] for record in records],
            [
                "SA_HEA_IMP",
                "SA_PHY_DIS",
                "SA_VIS_IMP",
                "SA_PSY_DIS",
                "SA_INT_DIS",
            ],
        )

    def test_affirmative_details_take_precedence_over_disability_code(self):
        rows = [{
            "PersonNumber": "1005",
            "Disability": "Yes",
            "DisabilityCode": "ORA_HRX_IN_DIS_J",
            "Description": "Synthetic test data: visual impairment.",
        }]
        lookups = self.LOOKUPS + [{
            "disability_code": "SA_VIS_IMP",
            "disability_name": "Visual impairment",
        }]

        records = build_person_disability_records(rows, lookups)

        self.assertEqual(records[0]["DisabilityCode"], "SA_VIS_IMP")

    def test_affirmative_unmatched_details_do_not_use_generic_code(self):
        rows = [{
            "PersonNumber": "1005",
            "Disability": "Yes",
            "Description": "Unknown disability details",
        }]

        with self.assertRaisesRegex(ValueError, "did not match an enabled"):
            build_person_disability_records(rows, self.LOOKUPS_WITH_DEFAULT)

    def test_does_not_create_file_without_disability_values(self):
        with tempfile.TemporaryDirectory() as output_dir:
            output_file = Path(output_dir) / "PersonDisability.dat"
            output_file.write_text("stale output", encoding="utf-8")

            result = generate_person_disability(
                [{"PersonNumber": "1005", "Disability": ""}],
                [],
                output_dir,
            )

            self.assertIsNone(result)
            self.assertFalse(output_file.exists())

    def test_writes_reference_metadata_and_merge_record(self):
        with tempfile.TemporaryDirectory() as output_dir:
            output_file = generate_person_disability(
                [{
                    "PersonNumber": "1005",
                    "Disability": "Hearing impairment - Deafness",
                    "StartDate": date(2025, 2, 12),
                    "LegislationCode": "SA",
                    "DisabilityDetails": "Partial hearing loss in left ear",
                    "Degree": 25,
                    "QuotaFTE": 1.0,
                    "WorkRestriction": "Requires phone amplification headset",
                    "Status": "A",
                }],
                self.LOOKUPS,
                output_dir,
            )

            self.assertEqual(output_file.name, "PersonDisability.dat")
            lines = output_file.read_text(encoding="utf-8").splitlines()
            self.assertEqual(
                lines[0],
                "METADATA|PersonDisability|"
                + "|".join(PERSON_DISABILITY_FIELDS),
            )
            self.assertEqual(
                lines[1],
                "MERGE|PersonDisability||ORA_HRX_IN_DIS_J|2025/02/12||SA||"
                "1005||||||||Partial hearing loss in left ear|25|1||"
                "Requires phone amplification headset|A",
            )

    def test_rejects_disability_not_in_database_lookup(self):
        with self.assertRaisesRegex(ValueError, "was not found"):
            build_person_disability_records(
                [{"PersonNumber": "1005", "Disability": "Unknown disability"}],
                self.LOOKUPS,
            )

    def test_requires_person_number_when_disability_is_populated(self):
        with self.assertRaisesRegex(ValueError, "Person Number is required"):
            build_person_disability_records(
                [{"Disability": "Hearing impairment - Deafness"}],
                self.LOOKUPS,
            )


if __name__ == "__main__":
    unittest.main()
