# -*- coding: utf-8 -*-
"""weeks/ 안의 편성표와 비중 파일을 최신 세트 기준으로 짝짓는다.

기존 weeks/ 루트에서는 동일 주차의 편성표 중 Git 이력상 가장 최근 파일을 사용한다.
weeks/{주차}/{버전}/ 안에서는 각 버전을 따로 처리한다.
비중 파일은 버전 폴더 또는 선택된 편성표와 같은 커밋을 기준으로 짝짓는다.

예:
    기존 편성표.xlsx
    기존 비중.xlsx
    새 편성표 (1).xlsx
    새 비중_새날짜.xlsx

새 편성표와 새 비중을 같은 커밋으로 올리면 새 세트만 사용하고 기존 세트는 무시한다.
버전 폴더는 서로 독립적으로 보존한다.
"""

import glob
import json
import os
import re
import subprocess

import openpyxl

import read_ratio_xlsx


SHEET = "(TV)편성기획"
VERSION_DIRS = {"v1", "v2", "confirmed", "final"}
WEEK_FOLDER = re.compile(r"\d{4}-\d{2}-\d+")
TITLE = re.compile(
    r"◆\s*(\d{4})년\s*(\d{2})월\s*(\d+)주\s*편성표\s*"
    r"\((\d{2})/(\d{2})\s*~\s*(\d{2})/(\d{2})\)"
)


def dup_suffix(path):
    """파일명 끝의 (n)을 숫자로 반환한다. 없으면 0."""

    base = os.path.splitext(os.path.basename(path))[0]
    match = re.search(r"\((\d+)\)\s*$", base)
    return int(match.group(1)) if match else 0


def classify(path):
    """파일 내용을 보고 편성표/비중을 판별한다."""

    try:
        workbook = openpyxl.load_workbook(path, read_only=True)
        sheet_names = workbook.sheetnames
        workbook.close()
    except Exception:
        return None

    if SHEET in sheet_names:
        return "sched"

    try:
        read_ratio_xlsx.read(path)
        return "ratio"
    except Exception:
        return None


def sched_key(path):
    """편성표 A1 제목에서 주차 정보를 읽는다."""

    workbook = openpyxl.load_workbook(path, read_only=True)
    title = str(workbook[SHEET]["A1"].value or "").strip()
    workbook.close()

    match = TITLE.match(title)
    if not match:
        return None, None

    year = int(match.group(1))
    month = int(match.group(2))
    week = int(match.group(3))
    start = f"{match.group(4)}/{match.group(5)}"
    end = f"{match.group(6)}/{match.group(7)}"
    cross = match.group(4) != match.group(6)

    return (
        f"{year}-{month}-{week}",
        {
            "y": year,
            "m": month,
            "w": week,
            "span": f"{start} ~ {end}",
            "cross": cross,
        },
    )


def version_context(path, weeks_dir):
    """Return a version id only for files in weeks/{week}/{version}/."""

    relative = os.path.relpath(path, weeks_dir)
    parts = relative.split(os.sep)
    if (
        len(parts) == 3
        and WEEK_FOLDER.fullmatch(parts[0])
        and parts[1] in VERSION_DIRS
    ):
        return parts[0], parts[1]
    return None, None


def git_info(path, root):
    """파일을 마지막으로 변경한 커밋 해시/시각/순서를 반환한다."""

    try:
        relative_path = os.path.relpath(path, root).replace(os.sep, "/")

        result = subprocess.run(
            [
                "git",
                "-C",
                root,
                "log",
                "-1",
                "--format=%H%x09%ct",
                "--",
                relative_path,
            ],
            capture_output=True,
            text=True,
            timeout=15,
            check=False,
        )

        value = result.stdout.strip()
        if not value:
            return None, 0, 0

        commit_hash, commit_time_text = value.split("\t", 1)

        rank_result = subprocess.run(
            [
                "git",
                "-C",
                root,
                "rev-list",
                "--count",
                commit_hash,
            ],
            capture_output=True,
            text=True,
            timeout=15,
            check=False,
        )

        rank_text = rank_result.stdout.strip()
        commit_rank = int(rank_text) if rank_text.isdigit() else 0

        return commit_hash, int(commit_time_text), commit_rank

    except Exception:
        return None, 0, 0


def freshness(item):
    """가장 최근 파일을 고르기 위한 정렬값."""

    # 로컬에서 아직 커밋하지 않은 새 파일은 기존 커밋 파일보다 최신으로 본다.
    if not item.get("commit"):
        return (
            2,
            item.get("mtime", 0),
            dup_suffix(item["file"]),
            os.path.basename(item["file"]),
        )

    return (
        1,
        item.get("commit_rank", 0),
        item.get("commit_time", 0),
        dup_suffix(item["file"]),
        os.path.basename(item["file"]),
    )


def load_saved_pairs(store):
    """기존 weeks/.pairs.json을 읽는다."""

    if not os.path.exists(store):
        return {}

    try:
        with open(store, encoding="utf-8") as file:
            value = json.load(file)
        return value if isinstance(value, dict) else {}
    except Exception:
        return {}


def select_latest_schedules(schedules):
    """동일 주차·버전 편성표 중 최신 파일 하나만 남긴다."""

    grouped = {}
    for schedule in schedules:
        group = schedule["group"]
        grouped.setdefault(group, []).append(schedule)

    selected = {}

    for key, candidates in grouped.items():
        chosen = max(candidates, key=freshness)
        selected[key] = chosen

        if len(candidates) > 1:
            print(
                f"  ↻ {chosen['key']} {chosen.get('version') or ''}: "
                "최신 편성표만 사용"
            )
            print(f"      사용: {os.path.basename(chosen['file'])}")

            for candidate in candidates:
                if candidate["file"] != chosen["file"]:
                    print(f"      무시: {os.path.basename(candidate['file'])}")

    return selected


def assign_ratio_key(
    ratio,
    all_schedules,
    selected_schedules,
    saved_pairs,
    weeks_dir,
):
    """비중 파일이 어느 주차·버전에 속하는지 결정한다."""

    relative_path = os.path.relpath(ratio["file"], weeks_dir)
    folder_key, folder_version = version_context(ratio["file"], weeks_dir)
    if folder_version:
        matching_groups = {
            schedule["group"]
            for schedule in selected_schedules.values()
            if schedule["key"] == folder_key
            and schedule.get("version") == folder_version
        }
        if len(matching_groups) == 1:
            return matching_groups.pop()
        return None

    # 가장 정확한 규칙: 같은 커밋에 올라온 편성표의 주차를 따른다.
    if ratio.get("commit"):
        same_commit = [
            schedule
            for schedule in all_schedules
            if schedule.get("commit") == ratio["commit"]
        ]
        same_keys = sorted({schedule["group"] for schedule in same_commit})

        if len(same_keys) == 1:
            return same_keys[0]

        if len(same_keys) > 1:
            print(
                f"  ! {os.path.basename(ratio['file'])}: "
                "같은 커밋에 서로 다른 주차 편성표가 여러 개라 "
                "자동 연결할 수 없음"
            )
            return None

    # 예전에 정상 연결되어 저장된 파일이면 그 연결을 재사용한다.
    saved_key = saved_pairs.get(relative_path)
    if saved_key in selected_schedules:
        return saved_key

    # Git 이력이 없는 로컬 실행에서는 수정 시각이 2분 이내인 편성표를 따른다.
    if not ratio.get("commit") and selected_schedules:
        nearest = min(
            selected_schedules.values(),
            key=lambda schedule: abs(schedule["mtime"] - ratio["mtime"]),
        )

        if abs(nearest["mtime"] - ratio["mtime"]) < 120:
            return nearest["key"]

    return None


def select_latest_ratio_set(schedule, ratio_records):
    """한 주차에 연결된 비중 파일 중 최신 세트만 선택한다."""

    if not ratio_records:
        return []

    schedule_commit = schedule.get("commit")

    # 최신 편성표와 같은 커밋에 올라온 비중이 최우선이다.
    same_commit = []
    if schedule_commit:
        same_commit = [
            ratio
            for ratio in ratio_records
            if ratio.get("commit") == schedule_commit
        ]

    if same_commit:
        selected = same_commit
    else:
        committed = [ratio for ratio in ratio_records if ratio.get("commit")]

        if committed:
            newest = max(committed, key=freshness)
            newest_commit = newest["commit"]
            selected = [
                ratio
                for ratio in committed
                if ratio["commit"] == newest_commit
            ]
        else:
            newest = max(ratio_records, key=freshness)
            selected = [
                ratio
                for ratio in ratio_records
                if abs(ratio["mtime"] - newest["mtime"]) < 120
            ]

    selected.sort(
        key=lambda ratio: (
            dup_suffix(ratio["file"]),
            os.path.basename(ratio["file"]),
        )
    )

    selected_paths = {ratio["file"] for ratio in selected}
    ignored = [
        ratio for ratio in ratio_records if ratio["file"] not in selected_paths
    ]

    if ignored:
        print(f"  ↻ {schedule['key']}: 이전 비중 {len(ignored)}개 무시")
        for ratio in ignored:
            print(f"      무시: {os.path.basename(ratio['file'])}")

    return selected


def pair(weeks_dir, root):
    """최신 편성표-비중 세트를 주차별로 반환한다."""

    files = [
        path
        for path in glob.glob(
            os.path.join(weeks_dir, "**", "*.xlsx"),
            recursive=True,
        )
        if not os.path.basename(path).startswith("~$")
    ]

    schedules = []
    ratios = []

    for path in files:
        file_type = classify(path)
        commit_hash, commit_time, commit_rank = git_info(path, root)
        folder_key, version = version_context(path, weeks_dir)

        common = {
            "file": path,
            "commit": commit_hash,
            "commit_time": commit_time,
            "commit_rank": commit_rank,
            "mtime": os.path.getmtime(path),
            "version": version,
        }

        if file_type == "sched":
            key, info = sched_key(path)

            if not key:
                print(
                    f"  ! {os.path.basename(path)}: "
                    "편성표 A1 제목을 읽지 못해 무시"
                )
                continue

            if version and folder_key != key:
                print(
                    f"  ! {os.path.basename(path)}: "
                    f"버전 폴더 주차({folder_key})와 A1 주차({key})가 달라 무시"
                )
                continue

            schedules.append(
                {
                    **common,
                    "key": key,
                    "group": f"{key}::{version}" if version else key,
                    "info": info,
                }
            )

        elif file_type == "ratio":
            ratios.append(common)

        else:
            print(
                f"  ! {os.path.basename(path)}: "
                "편성표도 비중도 아니므로 무시"
            )

    selected_schedules = select_latest_schedules(schedules)

    store = os.path.join(weeks_dir, ".pairs.json")
    saved_pairs = load_saved_pairs(store)

    by_group = {}
    for group, schedule in selected_schedules.items():
        key = schedule["key"]
        by_group[group] = {
            "file": schedule["file"],
            **schedule["info"],
            "key": key,
            "group": group,
            "version": schedule.get("version"),
            "ratios": [],
            "schedule_commit": schedule.get("commit"),
            "ratio_commits": [],
            "_schedule": schedule,
            "_ratio_records": [],
        }

    for ratio in ratios:
        group = assign_ratio_key(
            ratio=ratio,
            all_schedules=schedules,
            selected_schedules=selected_schedules,
            saved_pairs=saved_pairs,
            weeks_dir=weeks_dir,
        )

        if not group:
            print(
                f"  ! {os.path.basename(ratio['file'])}: "
                "어느 편성표의 비중인지 알 수 없어 무시"
            )
            continue

        if group not in by_group:
            print(
                f"  ! {os.path.basename(ratio['file'])}: "
                "배정된 주차·버전의 편성표가 없어 무시"
            )
            continue

        by_group[group]["_ratio_records"].append(ratio)

    new_saved_pairs = {}

    for group, item in by_group.items():
        selected_ratios = select_latest_ratio_set(
            schedule=item["_schedule"],
            ratio_records=item["_ratio_records"],
        )

        item["ratios"] = [ratio["file"] for ratio in selected_ratios]
        item["ratio_commits"] = sorted(
            {
                ratio["commit"]
                for ratio in selected_ratios
                if ratio.get("commit")
            }
        )
        revision_source = max(
            [item["_schedule"], *selected_ratios],
            key=freshness,
        )
        item["revision"] = revision_source.get("commit") or ""

        for ratio in selected_ratios:
            relative_path = os.path.relpath(ratio["file"], weeks_dir)
            new_saved_pairs[relative_path] = group

        del item["_schedule"]
        del item["_ratio_records"]

    try:
        with open(store, "w", encoding="utf-8") as file:
            json.dump(
                new_saved_pairs,
                file,
                ensure_ascii=False,
                indent=1,
            )
    except Exception as exception:
        print(f"  ! .pairs.json 저장 실패: {exception}")

    return sorted(
        by_group.values(),
        key=lambda item: (
            item["y"],
            item["m"],
            item["w"],
            item.get("version") or "",
        ),
    )


if __name__ == "__main__":
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    weeks_dir = os.path.join(root, "weeks")

    for item in pair(weeks_dir, root):
        ratio_names = " + ".join(
            os.path.basename(path) for path in item["ratios"]
        ) or "(비중 없음)"

        print(
            f"  {item['key']:10s} "
            f"{item['span']:15s} "
            f"{'월교차' if item['cross'] else '      '} "
            f"← {ratio_names}"
        )
