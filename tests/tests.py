#!/usr/bin/env python
# coding: utf-8

import unittest
import datetime
from icuschedulecheckup.main import is_absent

class TestIs_absent(unittest.TestCase):
    def get_schedule_patterns(self) -> dict:
        # Testing data
        prezence_den_v_tydnu, prezence_datum = {}, {}
        prezence_datum["Hry"] = [2, 3, 4, 5, 11, 12, 13]
        prezence_den_v_tydnu["Du"] = {
            "po_dopo": True,
            "po_odpo": False,
            "ut_dopo": True,
            "ut_odpo": True,
            "st_dopo": True,
            "st_odpo": True,
            "ct_dopo": True,
            "ct_odpo": False,
            "pa_dopo": True,
            "pa_odpo": True,
        }
        return {"day_of_week": prezence_den_v_tydnu, "date": prezence_datum}

    def test_day_of_week_absence(self):
        schedule_patterns = self.get_schedule_patterns()
        should_be_true = is_absent(
            name="Du",
            date=datetime.date(2025, 4, 17),
            part_of_day="odpo",
            schedule_patterns=schedule_patterns,
        )
        self.assertTrue(should_be_true)

        should_be_false = is_absent(
            name="Du",
            date=datetime.date(2025, 4, 17),
            part_of_day="dopo",
            schedule_patterns=schedule_patterns,
        )
        self.assertFalse(should_be_false)


    def test_date_absence(self):
        schedule_patterns = self.get_schedule_patterns()
        should_be_true = is_absent(
            name="Hry",
            date=datetime.date(2025, 4, 1),
            part_of_day="odpo",
            schedule_patterns=schedule_patterns,
        )
        self.assertTrue(should_be_true)

        should_be_false = is_absent(
            name="Hry",
            date=datetime.date(2025, 4, 2),
            part_of_day="dopo",
            schedule_patterns=schedule_patterns,
        )
        self.assertFalse(should_be_false)


if __name__ == "__main__":
    unittest.main()
