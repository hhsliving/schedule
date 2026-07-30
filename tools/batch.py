# -*- coding: utf-8 -*-
"""편성표 일괄 처리

폴더 구조 — 한 주차 = 한 폴더. 편성표와 비중 파일을 같이 넣으면 됩니다.

    weeks/
      2026-07-4/
        __2026년_07월_4주_편성표__0720___0726_.xlsx   ← 편성표 (파일명 자유)
        undefined_20260716.xlsx                       ← 비중 리포트 (파일명 자유, 없어도 됨)

  · 어느 쪽이 편성표이고 비중인지는 파일 내용을 보고 알아서 구분합니다.
  · 주차(연·월·주·기간·월교차)는 편성표 A1 제목에서 읽습니다. 폴더 이름은 참고용입니다.
  · 비중 파일이 없으면 비중표는 공란으로 만들어집니다.
"""

import glob
import json
import os
import re
import shutil
import subprocess
import sys

import openpyxl

import pair as pairing
import read_ratio_xlsx


# ---------------------------------------------------------------- 경로 및 설정

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WEEKS = os.environ.get("WEEKS_DIR", os.path.join(ROOT, "weeks"))
OUT = os.environ.get("IMG_DIR", os.path.join(ROOT, "images"))
WORK = os.path.dirname(os.path.abspath(__file__))
TMP = os.path.join(ROOT, "_batch")

SHEET = "(TV)편성기획"

# 전체 편성표의 열 너비를 어느 주차 파일 기준으로 맞출지 지정합니다.
# 직전 정상 너비였던 2026년 8월 2주차를 기본값으로 고정합니다.
WIDTH_TEMPLATE_KEY = os.environ.get(
    "WIDTH_TEMPLATE_KEY",
    "2026-08-2",
)

# 생성되는 PNG의 정상 가로 너비입니다.
EXPECTED_PNG_WIDTH = int(
    os.environ.get("EXPECTED_PNG_WIDTH", "1869")
)


# ---------------------------------------------------------------- 파일 분류

def classify(path):
    """편성표인지 비중 리포트인지 내용으로 판별.

    반환값:
        "sched" | "ratio" | None
    """

    try:
        wb = openpyxl.load_workbook(path, read_only=True)
        names = wb.sheetnames
        wb.close()
    except Exception:
        return None

    if SHEET in names:
        return "sched"

    try:
        read_ratio_xlsx.read(path)
        return "ratio"
    except Exception:
        return None


def scan():
    """weeks/의 파일들을 편성표-비중 세트로 짝지어 반환합니다."""

    items = pairing.pair(WEEKS, ROOT)

    # pair.py는 rkey 없이 반환하므로 key/rkey 정리
    for it in items:
        it["rkey"] = it["key"]
        it["key"] = (
            f"{it['y']}-{it['m']:02d}-{it['w']}"
        )

    return items


def get_ratio(it, quiet=True):
    """업로드된 비중 리포트에서 비중을 결정합니다.

    일반 주간:
        D열 현재주차 → 표의 "주"
        F열 월누적   → 표의 "월"

    월교차 주간:
        첫 번째 리포트 F열 → 지난달 누적
        두 번째 리포트 D열 → 새 달 1주차
    """

    rs = it["ratios"]

    if not rs:
        return None, None

    def warn(msg):
        if not quiet:
            print(f"  ! {it['key']}: {msg}")

    try:
        if it["cross"]:
            if len(rs) < 2:
                warn(
                    "월교차 주간인데 비중 리포트가 "
                    f"1개뿐입니다 ({os.path.basename(rs[0])}) "
                    "→ 일반 주간처럼 처리합니다"
                )

                week_value, month_value = (
                    read_ratio_xlsx.read(rs[0])
                )

                return (
                    week_value,
                    month_value,
                ), "리포트1(주의)"

            if len(rs) > 2:
                warn(
                    f"비중 리포트가 {len(rs)}개입니다 "
                    "→ 앞의 2개만 사용합니다"
                )

            _, month_prev = read_ratio_xlsx.read(rs[0])
            week_new, _ = read_ratio_xlsx.read(rs[1])

            return (
                week_new,
                month_prev,
            ), "리포트2(월교차)"

        if len(rs) > 1:
            warn(
                f"비중 리포트가 {len(rs)}개인데 "
                "월교차 주간이 아닙니다 "
                f"→ 첫 번째({os.path.basename(rs[0])})만 사용합니다"
            )

        week_value, month_value = read_ratio_xlsx.read(rs[0])

        return (
            week_value,
            month_value,
        ), "리포트"

    except Exception as exc:
        warn(f"비중 읽기 실패: {exc}")
        return None, None


# ---------------------------------------------------------------- 엑셀 생성

def build(it, template):
    """한 주차의 편성표 엑셀을 생성합니다."""

    key = it["key"]
    out = os.path.join(TMP, f"{key}.xlsx")

    env = dict(
        os.environ,
        SRCFILE=it["file"],
        OUTFILE=out,
        ROW1_PT="57",
        TEMPLATE=template,
    )

    ratio, _ = get_ratio(it)

    if ratio:
        vb, vc = ratio

        env["VB"] = json.dumps(
            vb,
            ensure_ascii=False,
        )
        env["VC"] = json.dumps(
            vc,
            ensure_ascii=False,
        )
        env.pop("BLANK", None)

    else:
        env["BLANK"] = "1"

    process = subprocess.run(
        [
            sys.executable,
            os.path.join(WORK, "build_sheet.py"),
        ],
        env=env,
        capture_output=True,
        text=True,
        cwd=WORK,
    )

    if not os.path.exists(out):
        error_text = process.stderr.strip()[-300:]
        print(f"  ! 실패 {key}: {error_text}")
        return None

    return out


def prep(xlsx, key):
    """렌더링 전 인쇄 설정을 적용합니다."""

    dst = os.path.join(TMP, f"r_{key}.xlsx")

    process = subprocess.run(
        [
            sys.executable,
            os.path.join(WORK, "render.py"),
        ],
        env=dict(
            os.environ,
            INFILE=xlsx,
            OUTFILE=dst,
        ),
        capture_output=True,
        text=True,
        cwd=WORK,
    )

    if not os.path.exists(dst):
        error_text = process.stderr.strip()[-200:]
        print(f"  ! 인쇄설정 실패 {key}: {error_text}")
        return None

    return dst


def find_width_template(items):
    """고정된 기준 주차의 편성표 파일을 찾습니다."""

    width_template = next(
        (
            item
            for item in items
            if item["key"] == WIDTH_TEMPLATE_KEY
        ),
        None,
    )

    if width_template is None:
        available = ", ".join(
            item["key"]
            for item in items
        )

        raise SystemExit(
            "\n열 너비 기준 편성표를 찾지 못했습니다.\n"
            f"기준 주차: {WIDTH_TEMPLATE_KEY}\n"
            f"현재 주차: {available}\n"
        )

    return width_template["file"]


def crop_and_validate_image(image_path, output_path, key, kind):
    """PDF 변환 이미지를 자르고 가로 너비를 검증합니다."""

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
        raise RuntimeError(
            "\n"
            "편성표 이미지 가로 너비가 정상값과 다릅니다.\n"
            f"파일: {kind}-{key}.png\n"
            f"현재 너비: {actual_width}px\n"
            f"정상 너비: {EXPECTED_PNG_WIDTH}px\n"
            f"높이: {actual_height}px\n"
            f"열 너비 기준 주차: {WIDTH_TEMPLATE_KEY}\n"
            "\n"
            "잘못된 너비의 이미지가 배포되는 것을 막기 위해 "
            "작업을 중단합니다."
        )

    cropped.save(
        output_path,
        optimize=True,
    )

    cropped.close()


# ---------------------------------------------------------------- 메인

def main():
    os.makedirs(OUT, exist_ok=True)
    os.makedirs(TMP, exist_ok=True)

    items = scan()

    if not items:
        raise SystemExit("처리할 주차가 없습니다")

    # 가장 최신 파일을 자동 기준으로 사용하지 않습니다.
    # 직전 정상 규격인 2026-08-2 편성표를 기준으로 고정합니다.
    default_template = find_width_template(items)

    # 필요할 때만 GitHub Actions 환경변수로 다른 파일을 지정할 수 있습니다.
    template = os.environ.get(
        "TEMPLATE_XLSX",
        default_template,
    )

    ratio_count = sum(
        1
        for item in items
        if get_ratio(item)[0]
    )

    print(
        f"대상 {len(items)}주차 / "
        f"비중 있음 {ratio_count}주차"
    )
    print(
        "열 너비 기준 주차: "
        f"{WIDTH_TEMPLATE_KEY}"
    )
    print(
        "열 너비 기준 파일: "
        f"{os.path.basename(template)}"
    )
    print(
        "PNG 정상 가로 너비: "
        f"{EXPECTED_PNG_WIDTH}px\n"
    )

    ready = []

    for it in items:
        ratio, source = get_ratio(
            it,
            quiet=False,
        )

        tag = (
            f"비중 {source}"
            if ratio
            else "비중 공란"
        )

        xlsx = build(
            it,
            template,
        )

        if not xlsx:
            continue

        rendered_xlsx = prep(
            xlsx,
            it["key"],
        )

        if not rendered_xlsx:
            continue

        ready.append(
            (
                it,
                rendered_xlsx,
            )
        )

        print(
            f"  {it['key']:<11} "
            f"{it['span']:<15} "
            f"{'월교차' if it['cross'] else '      '}  "
            f"{tag}"
        )

    print(
        f"\n엑셀 {len(ready)}/{len(items)} "
        "→ PDF/PNG 변환"
    )

    # 기존 PNG는 새 이미지가 모두 정상 생성된 뒤 교체하는 편이 안전하지만,
    # 기존 동작을 유지하기 위해 여기서 삭제합니다.
    for image_file in glob.glob(
        os.path.join(OUT, "*.png")
    ):
        os.remove(image_file)

    done = []

    for it, rendered_xlsx in ready:
        key = it["key"]
        pdf = rendered_xlsx.replace(
            ".xlsx",
            ".pdf",
        )

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
        )

        if not os.path.exists(pdf):
            print(f"  ! PDF 실패 {key}")
            continue

        ok = True

        for page, kind in (
            (1, "dept"),
            (2, "team"),
        ):
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

            generated = glob.glob(
                prefix + "-*.png"
            )

            if not generated:
                print(
                    f"  ! PNG 실패 {key} {kind}"
                )
                ok = False
                continue

            source_image = generated[0]
            output_image = os.path.join(
                OUT,
                f"{kind}-{key}.png",
            )

            try:
                crop_and_validate_image(
                    image_path=source_image,
                    output_path=output_image,
                    key=key,
                    kind=kind,
                )

            except Exception:
                if os.path.exists(output_image):
                    os.remove(output_image)

                raise

            finally:
                if os.path.exists(source_image):
                    os.remove(source_image)

        if ok:
            done.append(it)
            print(f"  ✓ {key}")

    print(
        f"\nPNG 완료 {len(done)}주차 × 2 "
        f"= {len(done) * 2}장"
    )

    # ----------------------------------------------------------------
    # index.html의 WEEKS 배열 갱신
    # ----------------------------------------------------------------

    done.sort(
        key=lambda item: (
            item["y"],
            item["m"],
            item["w"],
        )
    )

    lines = [
        (
            '  {{ y: {y}, m: {m}, w: {w}, '
            'span: "{span}", cross: {cross} }}'
        ).format(
            y=item["y"],
            m=item["m"],
            w=item["w"],
            span=item["span"],
            cross=(
                "true "
                if item["cross"]
                else "false"
            ),
        )
        for item in done
    ]

    weeks_array = (
        "const WEEKS = [\n"
        + ",\n".join(lines)
        + "\n];"
    )

    index_path = os.path.join(
        ROOT,
        "index.html",
    )

    with open(
        index_path,
        encoding="utf-8",
    ) as file:
        html = file.read()

    new_html, replace_count = re.subn(
        r"const WEEKS = \[.*?\];",
        lambda _: weeks_array,
        html,
        flags=re.S,
    )

    if replace_count == 0:
        raise SystemExit(
            "! index.html에서 WEEKS 배열을 찾지 못했습니다"
        )

    if new_html == html:
        print(
            f"index.html WEEKS {len(done)}주차 "
            "— 변경 없음"
        )

    else:
        with open(
            index_path,
            "w",
            encoding="utf-8",
        ) as file:
            file.write(new_html)

        print(
            f"index.html WEEKS {len(done)}주차 반영"
        )

    shutil.rmtree(
        TMP,
        ignore_errors=True,
    )


if __name__ == "__main__":
    main()
