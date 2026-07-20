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
import glob, os, re, subprocess, sys, json, shutil
import openpyxl
import read_ratio_xlsx
import pair as pairing

# ---- 경로 (환경변수로 덮어쓸 수 있음) ----
ROOT  = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WEEKS = os.environ.get("WEEKS_DIR", os.path.join(ROOT, "weeks"))
OUT   = os.environ.get("IMG_DIR",   os.path.join(ROOT, "images"))
WORK  = os.path.dirname(os.path.abspath(__file__))
TMP   = os.path.join(ROOT, "_batch")


# ---------------------------------------------------------------- 파일 분류
def classify(path):
    """편성표인지 비중 리포트인지 내용으로 판별 → 'sched' | 'ratio' | None"""
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
    """weeks/ 의 파일들을 편성표-비중 세트로 짝지어 반환"""
    items = pairing.pair(WEEKS, ROOT)
    # pair.py 는 rkey 없이 반환하므로 key/rkey 정리
    for it in items:
        it["rkey"] = it["key"]                       # 2026-7-5
        it["key"] = f"{it['y']}-{it['m']:02d}-{it['w']}"   # 2026-07-5 (이미지 파일명용)
    return items

def get_ratio(it, quiet=True):
    """비중 결정: 업로드된 비중 리포트만 사용

    일반 주간 (리포트 1개)
        D열 현재주차 → 표의 "주"
        F열 월누적   → 표의 "월"

    월교차 주간 (리포트 2개, 파일명 순)
        먼저 뽑은 리포트 = 지난달 월 비중  → 그 파일의 F열(월누적) → 표의 "OO월 누적"
        나중 뽑은 리포트 = 이번주 1주차 비중 → 그 파일의 D열(현재주차) → 표의 "△△월 1주차"
        (순서는 파일명 끝 (1) 접미사로 판단: 원본=전월, (1)=다음월)
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
                warn(f"월교차 주간인데 비중 리포트가 1개뿐입니다 "
                     f"({os.path.basename(rs[0])}) → 일반 주간처럼 처리합니다")
                w, m = read_ratio_xlsx.read(rs[0])
                return (w, m), "리포트1(주의)"
            if len(rs) > 2:
                warn(f"비중 리포트가 {len(rs)}개입니다 → 앞의 2개만 사용합니다")
            _, month_prev = read_ratio_xlsx.read(rs[0])   # 지난달 누적
            week_new, _   = read_ratio_xlsx.read(rs[1])   # 새 달 1주차
            return (week_new, month_prev), "리포트2(월교차)"
        else:
            if len(rs) > 1:
                warn(f"비중 리포트가 {len(rs)}개인데 월교차 주간이 아닙니다 "
                     f"→ 첫 번째({os.path.basename(rs[0])})만 사용합니다")
            w, m = read_ratio_xlsx.read(rs[0])
            return (w, m), "리포트"
    except Exception as e:
        warn(f"비중 읽기 실패: {e}")
        return None, None


# ---------------------------------------------------------------- 처리
def build(it, template):
    key = it["key"]
    out = os.path.join(TMP, f"{key}.xlsx")
    env = dict(os.environ, SRCFILE=it["file"], OUTFILE=out,
               ROW1_PT="57", TEMPLATE=template)
    r, _ = get_ratio(it)
    if r:
        vb, vc = r
        env["VB"] = json.dumps(vb, ensure_ascii=False)
        env["VC"] = json.dumps(vc, ensure_ascii=False)
        env.pop("BLANK", None)
    else:
        env["BLANK"] = "1"
    p = subprocess.run([sys.executable, os.path.join(WORK, "build_sheet.py")],
                       env=env, capture_output=True, text=True, cwd=WORK)
    if not os.path.exists(out):
        print(f"  ! 실패 {key}: {p.stderr.strip()[-300:]}")
        return None
    return out


def prep(xlsx, key):
    dst = os.path.join(TMP, f"r_{key}.xlsx")
    p = subprocess.run([sys.executable, os.path.join(WORK, "render.py")],
                       env=dict(os.environ, INFILE=xlsx, OUTFILE=dst),
                       capture_output=True, text=True, cwd=WORK)
    if not os.path.exists(dst):
        print(f"  ! 인쇄설정 실패 {key}: {p.stderr.strip()[-200:]}")
        return None
    return dst


def main():
    os.makedirs(OUT, exist_ok=True)
    os.makedirs(TMP, exist_ok=True)

    items = scan()
    if not items:
        raise SystemExit("처리할 주차가 없습니다")

    # 열 너비 기준 양식: 가장 최근 주차의 편성표 (과거 양식은 열이 좁아 여기 맞춰 통일)
    template = os.environ.get("TEMPLATE_XLSX", items[-1]["file"])

    n_ratio = sum(1 for i in items if get_ratio(i)[0])
    print(f"대상 {len(items)}주차 / 비중 있음 {n_ratio}주차")
    print(f"열 너비 기준 양식: {os.path.basename(template)}\n")

    ready = []
    for it in items:
        r, src = get_ratio(it, quiet=False)
        tag = f"비중 {src}" if r else "비중 공란"
        x = build(it, template)
        if not x:
            continue
        rx = prep(x, it["key"])
        if not rx:
            continue
        ready.append((it, rx))
        print(f"  {it['key']:<11} {it['span']:<15} "
              f"{'월교차' if it['cross'] else '      '}  {tag}")

    print(f"\n엑셀 {len(ready)}/{len(items)} → PDF/PNG 변환")

    from PIL import Image
    for f in glob.glob(os.path.join(OUT, "*.png")):
        os.remove(f)

    done = []
    for it, rx in ready:
        key = it["key"]
        pdf = rx.replace(".xlsx", ".pdf")
        if os.path.exists(pdf):
            os.remove(pdf)
        subprocess.run(["soffice", "--headless",
                        f"-env:UserInstallation=file:///tmp/lo_{key}",
                        "--convert-to", "pdf", rx, "--outdir", TMP],
                       capture_output=True, timeout=180)
        if not os.path.exists(pdf):
            print(f"  ! PDF 실패 {key}")
            continue
        ok = True
        for page, kind in ((1, "dept"), (2, "team")):
            pre = f"/tmp/p_{key}_{kind}"
            subprocess.run(["pdftoppm", "-png", "-r", "96",
                            "-f", str(page), "-l", str(page), pdf, pre], check=True)
            got = glob.glob(pre + "-*.png")
            if not got:
                print(f"  ! PNG 실패 {key} {kind}")
                ok = False
                continue
            im = Image.open(got[0]).convert("RGB")
            b = im.point(lambda x: 0 if x > 250 else 255).convert("L").getbbox()
            (im.crop(b) if b else im).save(os.path.join(OUT, f"{kind}-{key}.png"),
                                           optimize=True)
            os.remove(got[0])
        if ok:
            done.append(it)
            print(f"  ✓ {key}")

    print(f"\nPNG 완료 {len(done)}주차 × 2 = {len(done)*2}장")

    # ---- index.html 의 WEEKS 배열 갱신 ----
    done.sort(key=lambda k: (k["y"], k["m"], k["w"]))
    lines = ['  {{ y: {y}, m: {m}, w: {w}, span: "{span}", cross: {c} }}'.format(
        y=i["y"], m=i["m"], w=i["w"], span=i["span"],
        c="true " if i["cross"] else "false") for i in done]
    arr = "const WEEKS = [\n" + ",\n".join(lines) + "\n];"

    idx = os.path.join(ROOT, "index.html")
    html = open(idx, encoding="utf-8").read()
    new_html, n = re.subn(r"const WEEKS = \[.*?\];", lambda _: arr, html, flags=re.S)
    if n == 0:
        raise SystemExit("! index.html 에서 WEEKS 배열을 찾지 못했습니다")
    if new_html == html:
        print(f"index.html WEEKS {len(done)}주차 — 변경 없음")
    else:
        open(idx, "w", encoding="utf-8").write(new_html)
        print(f"index.html WEEKS {len(done)}주차 반영")

    shutil.rmtree(TMP, ignore_errors=True)


if __name__ == "__main__":
    main()
