# -*- coding: utf-8 -*-
"""
편성표 RAW 파일 일괄 처리
  1) A1 제목에서 연도/월/주차/기간/월교차 자동 판독
  2) 매크로 로직 적용 (비중은 기본 공란, RATIOS 에 있는 주차만 값 채움)
  3) 사업부/팀편성 각각 PNG 렌더
  4) 뷰어용 WEEKS 배열 출력
"""
import glob, os, re, subprocess, sys, json
import openpyxl

# 경로 (환경변수로 덮어쓸 수 있음)
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # 저장소 루트
UP   = os.environ.get("RAW_DIR", os.path.join(ROOT, "raw"))          # RAW 편성표 엑셀 폴더
WORK = os.path.dirname(os.path.abspath(__file__))                    # tools/
OUT  = os.environ.get("IMG_DIR", os.path.join(ROOT, "images"))
TMP  = os.path.join(ROOT, "_batch")

# 비중을 채울 주차만 여기에 (그 외 전부 공란)
RATIOS = {
    "2026-07-4": ({"가전": 0.0410, "리빙": 0.0713, "주방": 0.0674},
                  {"가전": 0.0449, "리빙": 0.0656, "주방": 0.0656}),
    "2026-07-5": ({"가전": 0.0358, "리빙": 0.0000, "주방": 0.1331},
                  {"가전": 0.0373, "리빙": 0.0644, "주방": 0.0666}),
}

# 열 너비 기준 양식 (7월 파일) — 과거 파일의 좁은 열을 이 양식으로 통일
# 열 너비 기준 양식 (7월 파일). 파일명이 바뀌면 이 값을 고쳐주세요.
TEMPLATE = os.environ.get("TEMPLATE_XLSX",
    os.path.join(UP, "__2026년_07월_4주_편성표__0720___0726___2_.xlsx"))

TITLE = re.compile(r"◆\s*(\d{4})년\s*(\d{2})월\s*(\d+)주\s*편성표\s*\((\d{2}/\d{2})\s*~\s*(\d{2}/\d{2})\)")

os.makedirs(OUT, exist_ok=True)
os.makedirs(TMP, exist_ok=True)

def scan():
    """업로드된 파일에서 주차 정보 읽기"""
    items = []
    for f in sorted(glob.glob(f"{UP}/*.xlsx")):
        wb = openpyxl.load_workbook(f, read_only=True)
        a1 = str(wb[wb.sheetnames[0]]["A1"].value or "").strip()
        wb.close()
        m = TITLE.match(a1)
        if not m:
            print(f"  ! 제목 형식 불일치, 건너뜀: {os.path.basename(f)}")
            continue
        y, mo, w = int(m.group(1)), int(m.group(2)), int(m.group(3))
        s, e = m.group(4), m.group(5)
        items.append({
            "file": f, "y": y, "m": mo, "w": w,
            "span": f"{s} ~ {e}",
            "cross": int(s.split("/")[0]) != int(e.split("/")[0]),
            "key": f"{y}-{mo:02d}-{w}",
        })
    items.sort(key=lambda k: (k["y"], k["m"], k["w"]))
    return items

def build(it):
    """매크로 적용 → xlsx"""
    key = it["key"]
    out = f"{TMP}/{key}.xlsx"
    env = dict(os.environ, SRCFILE=it["file"], OUTFILE=out, ROW1_PT="57", TEMPLATE=TEMPLATE)
    if key in RATIOS:
        vb, vc = RATIOS[key]
        env["VB"] = json.dumps(vb, ensure_ascii=False)
        env["VC"] = json.dumps(vc, ensure_ascii=False)
        env.pop("BLANK", None)
    else:
        env["BLANK"] = "1"
    r = subprocess.run([sys.executable, os.path.join(WORK, "build_sheet.py")],
                       env=env, capture_output=True, text=True, cwd=WORK)
    if not os.path.exists(out):
        print(f"  ! 실패 {key}: {r.stderr.strip()[-300:]}")
        return None
    return out

def prep(xlsx, key):
    """인쇄설정 적용 → 렌더용 xlsx"""
    r = subprocess.run([sys.executable, os.path.join(WORK, "render.py")],
                       env=dict(os.environ, INFILE=xlsx),
                       capture_output=True, text=True, cwd=WORK)
    src = f"{WORK}/render_exact.xlsx"
    if not os.path.exists(src):
        print(f"  ! 인쇄설정 실패 {key}: {r.stderr.strip()[-200:]}")
        return None
    dst = f"{TMP}/r_{key}.xlsx"
    os.replace(src, dst)
    return dst

if __name__ == "__main__":
    items = scan()
    print(f"대상 {len(items)}주차\n")

    ready = []
    for it in items:
        tag = "비중있음" if it["key"] in RATIOS else "비중공란"
        x = build(it)
        if not x:
            continue
        r = prep(x, it["key"])
        if not r:
            continue
        ready.append((it, r))
        print(f"  {it['key']:<11} {it['span']:<15} {'월교차' if it['cross'] else '      '}  {tag}")

    print(f"\n엑셀 생성 완료 {len(ready)}/{len(items)} → PDF/PNG 변환")

    from PIL import Image
    import shutil
    for f in glob.glob(f"{OUT}/*.png"):
        os.remove(f)

    done = []
    for it, rx in ready:
        key = it["key"]
        pdf = rx.replace(".xlsx", ".pdf")
        if os.path.exists(pdf): os.remove(pdf)
        subprocess.run(["soffice", "--headless",
                        f"-env:UserInstallation=file:///tmp/lo_{key}",
                        "--convert-to", "pdf", rx, "--outdir", TMP],
                       capture_output=True, timeout=180)
        if not os.path.exists(pdf):
            print(f"  ! PDF 실패 {key}"); continue
        ok = True
        for page, kind in ((1, "dept"), (2, "team")):
            pre = f"/tmp/p_{key}_{kind}"
            subprocess.run(["pdftoppm", "-png", "-r", "96",
                            "-f", str(page), "-l", str(page), pdf, pre], check=True)
            got = glob.glob(pre + "-*.png")
            if not got:
                print(f"  ! PNG 실패 {key} {kind}"); ok = False; continue
            im = Image.open(got[0]).convert("RGB")
            b = im.point(lambda x: 0 if x > 250 else 255).convert("L").getbbox()
            im = im.crop(b) if b else im
            im.save(f"{OUT}/{kind}-{key}.png", optimize=True)
            os.remove(got[0])
        if ok:
            done.append(it)
            print(f"  ✓ {key}")

    print(f"\nPNG 완료 {len(done)}주차 × 2 = {len(done)*2}장")
    with open(f"{TMP}/list.json", "w", encoding="utf-8") as f:
        json.dump([{k: v for k, v in it.items() if k != "file"} for it in done],
                  f, ensure_ascii=False, indent=1)
