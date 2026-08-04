# -*- coding: utf-8 -*-
"""편성표 이미지 생성.

GitHub에서 새 엑셀을 올렸을 때는 그 커밋에 포함된 주차만 다시 만든다.
기존 다른 주차 이미지는 삭제하거나 재생성하지 않는다.

수동 실행(workflow_dispatch) 또는 로컬 실행에서는 전체 주차를 다시 만든다.
"""

import glob
import json
import os
import re
import shutil
import subprocess
import sys

import pair as pairing
import read_ratio_xlsx


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WEEKS = os.environ.get("WEEKS_DIR", os.path.join(ROOT, "weeks"))
OUT = os.environ.get("IMG_DIR", os.path.join(ROOT, "images"))
WORK = os.path.dirname(os.path.abspath(__file__))
TMP = os.path.join(ROOT, "_batch")

WIDTH_TEMPLATE_KEY = os.environ.get("WIDTH_TEMPLATE_KEY", "2026-08-2")
EXPECTED_PNG_WIDTH = int(os.environ.get("EXPECTED_PNG_WIDTH", "1869"))
ZERO_SHA = "0000000000000000000000000000000000000000"


def scan():
    """weeks/의 파일을 최신 편성표-비중 세트로 정리한다."""

    items = pairing.pair(WEEKS, ROOT)

    for item in items:
        item["rkey"] = item["key"]
        item["key"] = f"{item['y']}-{item['m']:02d}-{item['w']}"

    return items


def get_ratio(item, quiet=True):
    """업로드된 비중 리포트에서 표에 넣을 값을 결정한다."""

    ratio_files = item["ratios"]

    if not ratio_files:
        return None, None

    def warn(message):
        if not quiet:
            print(f"  ! {item['key']}: {message}")

    try:
        if item["cross"]:
            if len(ratio_files) < 2:
                warn(
                    "월교차 주간인데 비중 리포트가 "
                    f"1개뿐입니다 ({os.path.basename(ratio_files[0])}) "
                    "→ 일반 주간처럼 처리합니다"
                )

                week_value, month_value = read_ratio_xlsx.read(ratio_files[0])
                return (week_value, month_value), "리포트1(주의)"

            if len(ratio_files) > 2:
                warn(
                    f"비중 리포트가 {len(ratio_files)}개입니다 "
                    "→ 앞의 2개만 사용합니다"
                )

            _, previous_month = read_ratio_xlsx.read(ratio_files[0])
            new_week, _ = read_ratio_xlsx.read(ratio_files[1])
            return (new_week, previous_month), "리포트2(월교차)"

        if len(ratio_files) > 1:
            warn(
                f"비중 리포트가 {len(ratio_files)}개인데 "
                "월교차 주간이 아닙니다 "
                f"→ 첫 번째({os.path.basename(ratio_files[0])})만 사용합니다"
            )

        week_value, month_value = read_ratio_xlsx.read(ratio_files[0])
        return (week_value, month_value), "리포트"

    except Exception as exception:
        warn(f"비중 읽기 실패: {exception}")
        return None, None


def build(item, template):
    """한 주차의 가공 엑셀을 생성한다."""

    key = item["key"]
    output = os.path.join(TMP, f"{key}.xlsx")

    environment = dict(
        os.environ,
        SRCFILE=item["file"],
        OUTFILE=output,
        ROW1_PT="57",
        TEMPLATE=template,
    )

    ratio, _ = get_ratio(item)

    if ratio:
        value_b, value_c = ratio
        environment["VB"] = json.dumps(value_b, ensure_ascii=False)
        environment["VC"] = json.dumps(value_c, ensure_ascii=False)
        environment.pop("BLANK", None)
    else:
        environment["BLANK"] = "1"

    process = subprocess.run(
        [sys.executable, os.path.join(WORK, "build_sheet.py")],
        env=environment,
        capture_output=True,
        text=True,
        cwd=WORK,
        check=False,
    )

    if not os.path.exists(output):
        error_text = process.stderr.strip()[-300:]
        print(f"  ! 실패 {key}: {error_text}")
        return None

    return output


def prep(xlsx, key):
    """렌더링 전 인쇄 설정을 적용한다."""

    output = os.path.join(TMP, f"r_{key}.xlsx")

    process = subprocess.run(
        [sys.executable, os.path.join(WORK, "render.py")],
        env=dict(os.environ, INFILE=xlsx, OUTFILE=output),
        capture_output=True,
        text=True,
        cwd=WORK,
        check=False,
    )

    if not os.path.exists(output):
        error_text = process.stderr.strip()[-200:]
        print(f"  ! 인쇄설정 실패 {key}: {error_text}")
        return None

    return output


def find_width_template(items):
    """고정된 기준 주차의 편성표 파일을 찾는다."""

    width_template = next(
        (item for item in items if item["key"] == WIDTH_TEMPLATE_KEY),
        None,
    )

    if width_template is None:
        available = ", ".join(item["key"] for item in items)
        raise SystemExit(
            "\n열 너비 기준 편성표를 찾지 못했습니다.\n"
            f"기준 주차: {WIDTH_TEMPLATE_KEY}\n"
            f"현재 주차: {available}\n"
        )

    return width_template["file"]


def crop_and_validate_image(image_path, output_path, key, kind):
    """PDF 변환 이미지를 자르고 가로 너비를 검증한다."""

    from PIL import Image

    with Image.open(image_path) as source:
        image = source.convert("RGB")

        bbox = image.point(
            lambda pixel: 0 if pixel > 250 else 255
        ).convert("L").getbbox()

        cropped = image.crop(bbox) if bbox else image.copy()

    actual_width = cropped.width
    actual_height = cropped.height

    if actual_width != EXPECTED_PNG_WIDTH:
        cropped.close()
        raise RuntimeError(
            "\n"
            "편성표 이미지 가로 너비가 정상값과 다릅니다.\n"
            f"파일: {kind}-{key}.png\n"
            f"현재 너비: {actual_width}px\n"
            f"정상 너비: {EXPECTED_PNG_WIDTH}px\n"
            f"높이: {actual_height}px\n"
            f"열 너비 기준 주차: {WIDTH_TEMPLATE_KEY}\n\n"
            "잘못된 너비의 이미지가 배포되는 것을 막기 위해 작업을 중단합니다."
        )

    cropped.save(output_path, optimize=True)
    cropped.close()


def commits_in_current_push():
    """이번 GitHub push에 포함된 커밋 해시를 반환한다."""

    current_sha = os.environ.get("GITHUB_SHA", "").strip()
    before_sha = os.environ.get("PUSH_BEFORE", "").strip()

    if not current_sha:
        return set()

    commits = {current_sha}

    if before_sha and before_sha != ZERO_SHA:
        process = subprocess.run(
            [
                "git",
                "-C",
                ROOT,
                "rev-list",
                f"{before_sha}..{current_sha}",
            ],
            capture_output=True,
            text=True,
            timeout=30,
            check=False,
        )

        found = {line.strip() for line in process.stdout.splitlines() if line.strip()}
        if found:
            commits = found

    return commits


def select_target_items(all_items):
    """이번 실행에서 실제로 이미지를 만들 주차만 선택한다."""

    full_rebuild = os.environ.get("FULL_REBUILD", "0") == "1"
    current_sha = os.environ.get("GITHUB_SHA", "").strip()

    # 로컬 실행 또는 수동 전체 실행은 전체를 처리한다.
    if full_rebuild or not current_sha:
        print("전체 재생성 모드")
        return list(all_items)

    push_commits = commits_in_current_push()
    targets = []

    for item in all_items:
        related_commits = set(item.get("ratio_commits", []))

        if item.get("schedule_commit"):
            related_commits.add(item["schedule_commit"])

        if related_commits & push_commits:
            targets.append(item)

    if targets:
        print(
            "이번 업로드 대상: "
            + ", ".join(item["key"] for item in targets)
        )
    else:
        print("이번 push에 새로 반영할 편성표 주차가 없습니다")

    return targets


def published_items(all_items):
    """사업부/팀 이미지가 모두 존재하는 주차만 반환한다."""

    result = []

    for item in all_items:
        key = item["key"]
        department_image = os.path.join(OUT, f"dept-{key}.png")
        team_image = os.path.join(OUT, f"team-{key}.png")

        if os.path.exists(department_image) and os.path.exists(team_image):
            result.append(item)
        else:
            print(f"  ! 이미지가 없는 주차는 목록에서 제외: {key}")

    return result


def render_item(item, rendered_xlsx):
    """한 주차의 PDF와 PNG 두 장을 만들고, 성공한 경우에만 기존 이미지를 교체한다."""

    key = item["key"]
    pdf = rendered_xlsx.replace(".xlsx", ".pdf")

    if os.path.exists(pdf):
        os.remove(pdf)

    subprocess.run(
        [
            "soffice",
            "--headless",
            f"-env:UserInstallation=file:///tmp/lo_{key}",
            "--convert-to",
            "pdf",
            rendered_xlsx,
            "--outdir",
            TMP,
        ],
        capture_output=True,
        timeout=180,
        check=False,
    )

    if not os.path.exists(pdf):
        print(f"  ! PDF 실패 {key}")
        return False

    temporary_outputs = {}

    try:
        for page, kind in ((1, "dept"), (2, "team")):
            prefix = f"/tmp/p_{key}_{kind}"

            subprocess.run(
                [
                    "pdftoppm",
                    "-png",
                    "-r",
                    "96",
                    "-f",
                    str(page),
                    "-l",
                    str(page),
                    pdf,
                    prefix,
                ],
                check=True,
            )

            generated = glob.glob(prefix + "-*.png")
            if not generated:
                raise RuntimeError(f"PNG 실패 {key} {kind}")

            source_image = generated[0]
            temporary_image = os.path.join(TMP, f"new-{kind}-{key}.png")

            if os.path.exists(temporary_image):
                os.remove(temporary_image)

            try:
                crop_and_validate_image(
                    image_path=source_image,
                    output_path=temporary_image,
                    key=key,
                    kind=kind,
                )
            finally:
                if os.path.exists(source_image):
                    os.remove(source_image)

            temporary_outputs[kind] = temporary_image

        # 두 장 모두 정상 생성된 뒤에만 기존 이미지를 교체한다.
        for kind in ("dept", "team"):
            final_image = os.path.join(OUT, f"{kind}-{key}.png")
            os.replace(temporary_outputs[kind], final_image)

        return True

    except Exception as exception:
        print(f"  ! 이미지 생성 실패 {key}: {exception}")

        for temporary_image in temporary_outputs.values():
            if os.path.exists(temporary_image):
                os.remove(temporary_image)

        return False


def update_index(all_items):
    """현재 실제 이미지가 존재하는 전체 주차로 index.html 목록을 갱신한다."""

    index_items = published_items(all_items)
    index_items.sort(key=lambda item: (item["y"], item["m"], item["w"]))

    lines = [
        (
            '  {{ y: {y}, m: {m}, w: {w}, '
            'span: "{span}", cross: {cross} }}'
        ).format(
            y=item["y"],
            m=item["m"],
            w=item["w"],
            span=item["span"],
            cross="true " if item["cross"] else "false",
        )
        for item in index_items
    ]

    weeks_array = "const WEEKS = [\n" + ",\n".join(lines) + "\n];"
    index_path = os.path.join(ROOT, "index.html")

    with open(index_path, encoding="utf-8") as file:
        html = file.read()

    new_html, replace_count = re.subn(
        r"const WEEKS = \[.*?\];",
        lambda _: weeks_array,
        html,
        flags=re.S,
    )

    if replace_count == 0:
        raise SystemExit("! index.html에서 WEEKS 배열을 찾지 못했습니다")

    if new_html == html:
        print(f"index.html WEEKS {len(index_items)}주차 — 변경 없음")
    else:
        with open(index_path, "w", encoding="utf-8") as file:
            file.write(new_html)
        print(f"index.html WEEKS {len(index_items)}주차 반영")


def main():
    os.makedirs(OUT, exist_ok=True)
    os.makedirs(TMP, exist_ok=True)

    all_items = scan()

    if not all_items:
        raise SystemExit("처리할 주차가 없습니다")

    target_items = select_target_items(all_items)
    default_template = find_width_template(all_items)
    template = os.environ.get("TEMPLATE_XLSX", default_template)

    ratio_count = sum(1 for item in target_items if get_ratio(item)[0])

    print(
        f"전체 {len(all_items)}주차 / "
        f"이번 생성 대상 {len(target_items)}주차 / "
        f"대상 중 비중 있음 {ratio_count}주차"
    )
    print(f"열 너비 기준 주차: {WIDTH_TEMPLATE_KEY}")
    print(f"열 너비 기준 파일: {os.path.basename(template)}")
    print(f"PNG 정상 가로 너비: {EXPECTED_PNG_WIDTH}px\n")

    ready = []

    for item in target_items:
        ratio, source = get_ratio(item, quiet=False)
        tag = f"비중 {source}" if ratio else "비중 공란"

        xlsx = build(item, template)
        if not xlsx:
            continue

        rendered_xlsx = prep(xlsx, item["key"])
        if not rendered_xlsx:
            continue

        ready.append((item, rendered_xlsx))

        print(
            f"  {item['key']:<11} "
            f"{item['span']:<15} "
            f"{'월교차' if item['cross'] else '      '}  "
            f"{tag}"
        )

    print(f"\n엑셀 {len(ready)}/{len(target_items)} → PDF/PNG 변환")

    done = []

    for item, rendered_xlsx in ready:
        if render_item(item, rendered_xlsx):
            done.append(item)
            print(f"  ✓ {item['key']}")

    print(f"\nPNG 완료 {len(done)}주차 × 2 = {len(done) * 2}장")

    update_index(all_items)
    shutil.rmtree(TMP, ignore_errors=True)


if __name__ == "__main__":
    main()
