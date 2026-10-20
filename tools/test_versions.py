import tempfile
import os
import sys
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

sys.path.insert(0, os.path.dirname(__file__))

import batch
import pair


class VersionPairingTests(unittest.TestCase):
    def test_legacy_pair_mapping_keeps_its_original_key_format(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            weeks = root / "weeks"
            weeks.mkdir()
            (weeks / "schedule.xlsx").write_bytes(b"schedule")
            (weeks / "ratio.xlsx").write_bytes(b"ratio")

            def classify(path):
                return "ratio" if "ratio" in Path(path).name else "sched"

            with (
                patch.object(pair, "classify", side_effect=classify),
                patch.object(
                    pair,
                    "sched_key",
                    return_value=(
                        "2026-10-4",
                        {
                            "y": 2026,
                            "m": 10,
                            "w": 4,
                            "span": "10/19 ~ 10/25",
                            "cross": False,
                        },
                    ),
                ),
                patch.object(
                    pair,
                    "git_info",
                    return_value=("same-commit", 1, 1),
                ),
            ):
                items = pair.pair(str(weeks), str(root))

            saved_pairs = (weeks / ".pairs.json").read_text(encoding="utf-8")

        self.assertEqual(len(items), 1)
        self.assertIsNone(items[0]["version"])
        self.assertIn('"ratio.xlsx": "2026-10-4"', saved_pairs)

    def test_version_directories_keep_schedule_and_ratio_sets_separate(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            weeks = root / "weeks"
            for version in ("v1", "v2"):
                folder = weeks / "2026-10-5" / version
                folder.mkdir(parents=True)
                (folder / "schedule.xlsx").write_bytes(b"schedule")
                (folder / "ratio.xlsx").write_bytes(b"ratio")

            def classify(path):
                return "ratio" if "ratio" in Path(path).name else "sched"

            def schedule_key(_path):
                return (
                    "2026-10-5",
                    {
                        "y": 2026,
                        "m": 10,
                        "w": 5,
                        "span": "10/26 ~ 11/01",
                        "cross": True,
                    },
                )

            def git_info(path, _root):
                version = Path(path).parent.name
                return f"{version}-commit", 1, 1

            with (
                patch.object(pair, "classify", side_effect=classify),
                patch.object(pair, "sched_key", side_effect=schedule_key),
                patch.object(pair, "git_info", side_effect=git_info),
            ):
                items = pair.pair(str(weeks), str(root))

        by_version = {item["version"]: item for item in items}
        self.assertEqual(set(by_version), {"v1", "v2"})
        for version, item in by_version.items():
            self.assertEqual(
                [Path(path).parent.name for path in item["ratios"]],
                [version],
            )
            self.assertEqual(item["key"], "2026-10-5")
            self.assertEqual(item["revision"], f"{version}-commit")


class VersionIndexTests(unittest.TestCase):
    def test_version_labels_match_teams_stage_names(self):
        self.assertEqual(batch.VERSION_LABELS["v1"], "1차 가안")
        self.assertEqual(batch.VERSION_LABELS["v2"], "2차 가안")
        self.assertEqual(batch.VERSION_LABELS["confirmed"], "확정안")
        self.assertEqual(batch.VERSION_LABELS["final"], "최종안")

    def test_deletion_commit_is_detected_without_render_targets(self):
        with (
            patch.object(
                batch,
                "commits_in_current_push",
                return_value={"delete-commit"},
            ),
            patch.object(
                batch.subprocess,
                "run",
                return_value=Mock(stdout="편성표 버전 삭제: 2026-10-5 1차 가안\n"),
            ),
        ):
            self.assertTrue(batch.has_version_deletion_commit())

    def test_deleting_a_version_prunes_only_its_rendered_files(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            images = root / "images"
            pdfs = root / "pdfs"
            images.mkdir()
            pdfs.mkdir()
            for name in (
                "dept-2026-10-5--v1.png",
                "team-2026-10-5--v1.png",
                "dept-2026-10-5--v2.png",
                "team-2026-10-5--v2.png",
            ):
                (images / name).touch()
            for name in (
                "schedule-2026-10-5--v1.pdf",
                "schedule-2026-10-5--v2.pdf",
            ):
                (pdfs / name).touch()

            active_items = [
                {
                    "artifact_key": "2026-10-5--v2",
                    "version": "v2",
                }
            ]
            with (
                patch.object(batch, "OUT", str(images)),
                patch.object(batch, "PDF_OUT", str(pdfs)),
            ):
                batch.prune_orphaned_version_artifacts(active_items)

            self.assertFalse((images / "dept-2026-10-5--v1.png").exists())
            self.assertFalse((images / "team-2026-10-5--v1.png").exists())
            self.assertFalse((pdfs / "schedule-2026-10-5--v1.pdf").exists())
            self.assertTrue((images / "dept-2026-10-5--v2.png").exists())
            self.assertTrue((pdfs / "schedule-2026-10-5--v2.pdf").exists())

    def test_index_lists_only_rendered_versions_and_keeps_legacy_weeks(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            images = root / "images"
            pdfs = root / "pdfs"
            images.mkdir()
            pdfs.mkdir()
            (root / "index.html").write_text(
                "const MONTHLY_RATIOS = [];\nconst WEEKS = [];\n",
                encoding="utf-8",
            )

            items = [
                {
                    "key": "2026-10-4",
                    "artifact_key": "2026-10-4",
                    "y": 2026,
                    "m": 10,
                    "w": 4,
                    "span": "10/19 ~ 10/25",
                    "cross": False,
                    "version": None,
                    "ratios": [],
                },
                {
                    "key": "2026-10-5",
                    "artifact_key": "2026-10-5--v1",
                    "y": 2026,
                    "m": 10,
                    "w": 5,
                    "span": "10/26 ~ 11/01",
                    "cross": True,
                    "version": "v1",
                    "schedule_commit": "revision-v1",
                    "ratio_commits": [],
                    "ratios": [],
                },
                {
                    "key": "2026-10-5",
                    "artifact_key": "2026-10-5--v2",
                    "y": 2026,
                    "m": 10,
                    "w": 5,
                    "span": "10/26 ~ 11/01",
                    "cross": True,
                    "version": "v2",
                    "schedule_commit": "revision-v2",
                    "ratio_commits": [],
                    "ratios": [],
                },
            ]
            for item in items:
                (images / f"dept-{item['artifact_key']}.png").touch()
                (images / f"team-{item['artifact_key']}.png").touch()

            with (
                patch.object(batch, "ROOT", str(root)),
                patch.object(batch, "OUT", str(images)),
                patch.object(batch, "PDF_OUT", str(pdfs)),
            ):
                batch.update_index(items, target_items=[])

            html = (root / "index.html").read_text(encoding="utf-8")

        self.assertIn(
            '{ y: 2026, m: 10, w: 4, span: "10/19 ~ 10/25", cross: false }',
            html,
        )
        self.assertIn(
            '"revision": "revision-v1"',
            html,
        )
        self.assertIn(
            '"revision": "revision-v2"',
            html,
        )
        self.assertEqual(html.count("versions: "), 1)
        self.assertLess(html.index('"id": "v1"'), html.index('"id": "v2"'))


if __name__ == "__main__":
    unittest.main()
