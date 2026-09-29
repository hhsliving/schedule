import os
import sys
import unittest
from unittest.mock import patch

sys.path.insert(0, os.path.dirname(__file__))

import batch


def item(year, month, week, span, cross=False, ratios=None):
    return {
        "y": year,
        "m": month,
        "w": week,
        "span": span,
        "cross": cross,
        "ratios": ratios or [],
    }


class MonthlyRatioTests(unittest.TestCase):
    def test_month_end_week_is_found_across_month_boundary(self):
        value = item(2026, 1, 5, "01/26 ~ 02/01", cross=True)

        self.assertEqual(
            batch.month_keys_in_span(value),
            (
                {"2026-01", "2026-02"},
                {"2026-01"},
            ),
        )

    def test_year_boundary_uses_previous_year_for_december(self):
        value = item(2027, 1, 1, "12/28 ~ 01/03", cross=True)

        self.assertEqual(
            batch.month_keys_in_span(value),
            (
                {"2026-12", "2027-01"},
                {"2026-12"},
            ),
        )

    def test_month_end_cross_week_uses_report_for_ending_month(self):
        value = item(
            2026,
            1,
            5,
            "01/26 ~ 02/01",
            cross=True,
            ratios=["january.xlsx", "february.xlsx"],
        )
        january_values = {
            "가전": 0.1,
            "리빙": 0.2,
            "주방": 0.3,
        }
        with patch.object(
            batch.read_ratio_xlsx,
            "read",
            return_value=({}, january_values),
        ) as read:
            result = batch.monthly_values_for_item(value, 2026, 1)

        read.assert_called_once_with("january.xlsx")
        self.assertEqual(result["사업부"], 0.6)
        self.assertEqual(result["가전"], 0.1)

    def test_incremental_update_changes_only_month_whose_closing_week_was_uploaded(self):
        january_closing = item(
            2026,
            1,
            5,
            "01/26 ~ 02/01",
            cross=True,
            ratios=["january.xlsx", "february-first-week.xlsx"],
        )
        february_closing = item(
            2026,
            2,
            4,
            "02/23 ~ 03/01",
            cross=True,
            ratios=["february.xlsx", "march-first-week.xlsx"],
        )
        previous = [
            {
                "y": 2026,
                "m": 1,
                "closing_week": "2026-01-5",
                "span": "01/26 ~ 02/01",
                "values": {"사업부": 0.5, "가전": 0.1, "리빙": 0.2, "주방": 0.2},
            },
            {
                "y": 2026,
                "m": 2,
                "closing_week": "2026-02-4",
                "span": "02/23 ~ 03/01",
                "values": {"사업부": 0.4, "가전": 0.1, "리빙": 0.1, "주방": 0.2},
            },
        ]
        report_values = {
            "가전": 0.12,
            "리빙": 0.23,
            "주방": 0.34,
        }

        with patch.object(
            batch.read_ratio_xlsx,
            "read",
            return_value=({}, report_values),
        ) as read:
            result = batch.monthly_ratio_snapshot(
                [january_closing, february_closing],
                previous,
                [february_closing],
            )

        self.assertEqual([row["m"] for row in result], [1, 2])
        self.assertEqual(result[0]["values"]["사업부"], 0.5)
        self.assertEqual(result[1]["values"]["가전"], 0.12)
        read.assert_called_once_with("february.xlsx")

    def test_initial_snapshot_keeps_years_as_distinct_month_keys(self):
        older_month = item(2026, 1, 5, "01/26 ~ 02/01", cross=True)
        newer_month = item(2027, 1, 5, "01/25 ~ 01/31")
        values = {"가전": 0.1, "리빙": 0.2, "주방": 0.3}
        with patch.object(
            batch.read_ratio_xlsx,
            "read",
            return_value=({}, values),
        ):
            result = batch.monthly_ratio_snapshot(
                [older_month, newer_month],
                [],
                [],
            )

        self.assertEqual(
            [(row["y"], row["m"]) for row in result],
            [(2026, 1), (2026, 2), (2027, 1)],
        )

    def test_new_year_month_end_update_preserves_previous_year(self):
        older_month = item(
            2026,
            1,
            5,
            "01/26 ~ 02/01",
            cross=True,
            ratios=["old-january.xlsx", "old-february.xlsx"],
        )
        newer_month = item(
            2027,
            1,
            5,
            "01/25 ~ 01/31",
            ratios=["new-january.xlsx"],
        )
        previous = [
            {
                "y": 2026,
                "m": 1,
                "closing_week": "2026-01-5",
                "span": "01/26 ~ 02/01",
                "values": {"사업부": 0.5, "가전": 0.1, "리빙": 0.2, "주방": 0.2},
            }
        ]
        new_values = {"가전": 0.11, "리빙": 0.22, "주방": 0.33}
        with patch.object(
            batch.read_ratio_xlsx,
            "read",
            return_value=({}, new_values),
        ):
            result = batch.monthly_ratio_snapshot(
                [older_month, newer_month],
                previous,
                [newer_month],
            )

        self.assertEqual(
            [(row["y"], row["m"]) for row in result],
            [(2026, 1), (2027, 1)],
        )
        self.assertEqual(result[0]["values"]["가전"], 0.1)
        self.assertEqual(result[1]["values"]["가전"], 0.11)


if __name__ == "__main__":
    unittest.main()
