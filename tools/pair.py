# -*- coding: utf-8 -*-
"""weeks/ 안에 다 섞인 파일들을 편성표-비중 세트로 짝짓는다.

원칙: "같은 커밋에 올라온 편성표와 비중은 한 세트"
  - 편성표는 A1 제목으로 주차를 안다.
  - 비중은 자기가 어느 주차인지 모른다 → 같은 커밋의 편성표를 따라간다.
  - git 이력이 있으면 커밋 단위로 묶고, 없으면(로컬 최초 실행 등)
    파일 mtime 이 비슷한 것끼리 묶는다.

주차별 짝 정보는 weeks/.pairs.json 에 저장해 다음 실행 때 재사용한다.
(한 번 짝지어진 비중은 이후 커밋에서 편성표와 안 올라와도 유지된다.)
"""
import os, re, json, glob, subprocess
import openpyxl
import read_ratio_xlsx

SHEET = "(TV)편성기획"
TITLE = re.compile(r"◆\s*(\d{4})년\s*(\d{2})월\s*(\d+)주\s*편성표\s*\((\d{2})/(\d{2})\s*~\s*(\d{2})/(\d{2})\)")


def dup_suffix(path):
    """파일명 끝 (n) → n, 없으면 0. 방송기간 (0713 ~ 0719) 는 오인 안 함."""
    base = os.path.splitext(os.path.basename(path))[0]
    m = re.search(r"\((\d+)\)\s*$", base)
    return int(m.group(1)) if m else 0


def classify(path):
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


def sched_key(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    a1 = str(wb[SHEET]["A1"].value or "").strip()
    wb.close()
    m = TITLE.match(a1)
    if not m:
        return None, None
    y, mo, w = int(m.group(1)), int(m.group(2)), int(m.group(3))
    s, e = f"{m.group(4)}/{m.group(5)}", f"{m.group(6)}/{m.group(7)}"
    cross = m.group(4) != m.group(6)
    return f"{y}-{mo}-{w}", {"y": y, "m": mo, "w": w, "span": f"{s} ~ {e}", "cross": cross}


def commit_of(path, root):
    """이 파일이 마지막으로 '추가/변경된' 커밋 해시 (git 없으면 None)"""
    try:
        rel = os.path.relpath(path, root)
        out = subprocess.run(["git", "-C", root, "log", "-1", "--format=%H", "--", rel],
                             capture_output=True, text=True, timeout=15)
        h = out.stdout.strip()
        return h or None
    except Exception:
        return None


def pair(weeks_dir, root):
    """→ [{key, y, m, w, span, cross, sched, ratios:[...]}]  주차별로"""
    files = [f for f in glob.glob(os.path.join(weeks_dir, "*.xlsx"))
             if not os.path.basename(f).startswith("~$")]

    scheds, ratios = [], []
    for f in files:
        k = classify(f)
        if k == "sched":
            key, info = sched_key(f)
            if key:
                scheds.append({"file": f, "key": key, "info": info,
                               "commit": commit_of(f, root),
                               "mtime": os.path.getmtime(f)})
            else:
                print(f"  ! {os.path.basename(f)}: 편성표인데 A1 제목을 못 읽음 → 무시")
        elif k == "ratio":
            ratios.append({"file": f, "commit": commit_of(f, root),
                           "mtime": os.path.getmtime(f)})
        else:
            print(f"  ! {os.path.basename(f)}: 편성표도 비중도 아님 → 무시")

    # 저장된 짝 정보 불러오기 (비중파일 경로 → 주차키)
    store = os.path.join(weeks_dir, ".pairs.json")
    saved = {}
    if os.path.exists(store):
        try:
            saved = json.load(open(store, encoding="utf-8"))
        except Exception:
            saved = {}

    # 비중 → 주차 배정
    def assign(r):
        rel = os.path.relpath(r["file"], weeks_dir)
        # 1) 이미 저장돼 있으면 그대로
        if rel in saved:
            return saved[rel]
        # 2) 같은 커밋의 편성표
        if r["commit"]:
            same = [s for s in scheds if s["commit"] == r["commit"]]
            if len(same) == 1:
                return same[0]["key"]
            if len(same) > 1:
                print(f"  ! {os.path.basename(r['file'])}: 같은 커밋에 편성표가 "
                      f"{len(same)}개 → mtime 으로 가장 가까운 것 선택")
                return min(same, key=lambda s: abs(s["mtime"] - r["mtime"]))["key"]
        # 3) git 없음: mtime 이 가장 가까운 편성표
        if scheds:
            near = min(scheds, key=lambda s: abs(s["mtime"] - r["mtime"]))
            if abs(near["mtime"] - r["mtime"]) < 120:   # 2분 이내면 같은 세트로 간주
                return near["key"]
        return None

    by_key = {s["key"]: {"file": s["file"], **s["info"], "key": s["key"], "ratios": []}
              for s in scheds}

    new_saved = dict(saved)
    for r in ratios:
        key = assign(r)
        if not key:
            print(f"  ! {os.path.basename(r['file'])}: 어느 편성표와 세트인지 알 수 없음 "
                  f"(편성표와 같이 커밋했는지 확인) → 건너뜀")
            continue
        if key not in by_key:
            print(f"  ! {os.path.basename(r['file'])}: 배정된 주차 {key} 의 편성표가 없음 → 건너뜀")
            continue
        by_key[key]["ratios"].append(r["file"])
        new_saved[os.path.relpath(r["file"], weeks_dir)] = key

    # 비중 순서: 접미사 번호 (원본 0=전월, (1)=다음월)
    for v in by_key.values():
        v["ratios"].sort(key=dup_suffix)

    # 짝 정보 저장 (다음 실행 때 재사용)
    try:
        json.dump(new_saved, open(store, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    except Exception:
        pass

    return sorted(by_key.values(), key=lambda x: (x["y"], x["m"], x["w"]))


if __name__ == "__main__":
    import sys
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    wk = os.path.join(root, "weeks")
    for it in pair(wk, root):
        rs = " + ".join(os.path.basename(r) for r in it["ratios"]) or "(비중 없음)"
        print(f"  {it['key']:10s} {it['span']:15s} {'월교차' if it['cross'] else '     '}  ← {rs}")
