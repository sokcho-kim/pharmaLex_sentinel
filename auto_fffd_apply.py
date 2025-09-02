# -*- coding: utf-8 -*-
"""
U+FFFD(�) 자동 교정기
- 입력:
  C:\Jimin\pharmaLex_sentinel\out\mapping_candidates.csv   # build_mapping_from_pdf.py 결과
  C:\Jimin\pharmaLex_sentinel\data\요양심사약제_후처리.xlsx # 원본 엑셀(시트 1개/여러개 모두 OK)
- 처리:
  1) mapping_candidates의 점수로 best_candidate 자동 확정
     - 기준: top_total >= MIN_HITS and top_total >= MARGIN_RATIO * second_total
  2) 점수 애매하면 문맥 규칙(숫자·단위·그리스문자)로 보정
  3) 해당 셀의 '�'만 교체 (다른 문자는 건드리지 않음)
- 출력:
  out/요양심사약제_후처리_fffd_autofixed.xlsx
  out/fffd_autofix_log.csv  (어디를 무엇으로 왜 바꿨는지)
"""

import os, re, pandas as pd
from collections import Counter

BASE = r"C:\Jimin\pharmaLex_sentinel"
IN_CAND = os.path.join(BASE, r"out\mapping_candidates.csv")
IN_XLSX = os.path.join(BASE, r"out\요양심사약제_후처리_fffd_autofixed.xlsx")  # 이미 처리된 파일 사용
OUT_DIR = os.path.join(BASE, "out")
OUT_XLSX = os.path.join(OUT_DIR, "요양심사약제_후처리_fffd_autofixed_v2.xlsx")
OUT_LOG  = os.path.join(OUT_DIR, "fffd_autofix_log_v2.csv")

# --------- 자동 확정 기준 (형님 원하는대로 '확신만' 자동) ----------
MIN_HITS = 3         # top 후보 최소 히트수
MARGIN_RATIO = 2.0   # top >= ratio * second 이면 자동 확정
# ------------------------------------------------------------

# 애매할 때 쓰는 문맥 규칙(보수적)
NUM = r"(?:\d+(?:\.\d+)?)"
HEURISTICS = [
    # 숫자 � g  → 숫자 ㎍ g 로 많이 깨짐 (예: 700�g, 10�g)
    (re.compile(rf"({NUM})\s*�\s*g", re.IGNORECASE), r"\1 ㎍ g"),
    # 숫자 � m l → ㎖ (예: 5 � m l 형태로 쪼개진 케이스 방어)
    (re.compile(rf"({NUM})\s*�\s*m\s*l", re.IGNORECASE), r"\1 ㎖"),
    # 숫자 � l  → ㎖ (예: 5 �l)
    (re.compile(rf"({NUM})\s*�l", re.IGNORECASE), r"\1 ㎖"),
    # a- / b- / g- 앞의 � → α/β/γ 추정
    (re.compile(r"\b�-?\s*blocker", re.IGNORECASE), "α-blocker"),
    (re.compile(r"\b�-?\s*interferon", re.IGNORECASE), "α-interferon"),
    (re.compile(r"\bpeginterferon\s+�-?1", re.IGNORECASE), "peginterferon α-1"),
]

def parse_scores(scores_str: str):
    # "㎍:12(p459|p461) | ㎎:3(p21) | ㎖:0" → [('㎍',12), ('㎎',3), ('㎖',0)]
    if not isinstance(scores_str, str) or not scores_str.strip():
        return []
    out = []
    for part in [p.strip() for p in scores_str.split("|")]:
        m = re.match(r"^(.+?):\s*(\d+)", part)
        if m:
            out.append((m.group(1).strip(), int(m.group(2))))
    return out

def confident_choice(best: str, scores_str: str):
    scores = parse_scores(scores_str)
    if not scores:
        return False, ""
    scores.sort(key=lambda x: -x[1])
    top_cand, top_total = scores[0]
    second_total = scores[1][1] if len(scores) >= 2 else 0
    if top_cand != best:
        # best_candidate는 이미 점수 기반이므로 그대로 신뢰
        pass
    if top_total >= MIN_HITS and (second_total == 0 or top_total >= MARGIN_RATIO * second_total):
        return True, best
    return False, ""

def apply_heuristics(text: str):
    new = text
    applied = []
    for pat, repl in HEURISTICS:
        new2, n = pat.subn(repl, new)
        if n > 0:
            applied.append(f"{pat.pattern} -> {repl} x{n}")
        new = new2
    return new, applied

def load_candidates():
    df = pd.read_csv(IN_CAND, dtype=str).fillna("")
    # (sheet,row,column) → (value, best, scores)
    rows = {}
    for _, r in df.iterrows():
        key = (r["sheet"], str(r["row"]), r["column"])
        rows[key] = {
            "value": r["value"],
            "best": r.get("best_candidate",""),
            "scores": r.get("candidate_scores","")
        }
    return rows

def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    cand = load_candidates()

    xls = pd.ExcelFile(IN_XLSX)
    writer = pd.ExcelWriter(OUT_XLSX, engine="openpyxl")
    logs = []

    for sheet in xls.sheet_names:
        df = pd.read_excel(IN_XLSX, sheet_name=sheet, dtype=str)
        for r in range(len(df)):
            for c, col in enumerate(df.columns):
                val = df.iat[r, c]
                if pd.isna(val) or "�" not in str(val):
                    continue
                s0 = str(val)
                s = s0

                # 1) mapping_candidates 기반 자동 확정 시도
                key = (sheet, str(r+2), col)  # 엑셀 표시행 기준(row+2)
                applied_reason = ""

                if key in cand:
                    ok, choice = confident_choice(cand[key]["best"], cand[key]["scores"])
                    if ok and choice:
                        s = s.replace("�", choice)
                        applied_reason = f"auto-best:{choice}"
                
                # 2) 점수 애매했거나 후보표에 없으면 문맥 휴리스틱
                if "�" in s:
                    s_heur, heur_applied = apply_heuristics(s)
                    if s_heur != s:
                        s = s_heur
                        if applied_reason:
                            applied_reason += " + heuristics"
                        else:
                            applied_reason = "heuristics"

                # 3) 그래도 남아있으면 최후의 안전장치(치환 안 함)
                if s != s0:
                    df.iat[r, c] = s
                    logs.append({
                        "sheet": sheet,
                        "row": r+2,
                        "column": col,
                        "before": s0,
                        "after": s,
                        "reason": applied_reason if applied_reason else "n/a"
                    })

        df.to_excel(writer, sheet_name=sheet, index=False)
    writer.close()

    pd.DataFrame(logs).to_csv(OUT_LOG, index=False, encoding="utf-8-sig")
    print("[OK] 엑셀 저장:", OUT_XLSX)
    print("[OK] 로그 저장 :", OUT_LOG)
    print(f"[INFO] 총 변경 셀 수: {len(logs)}")

if __name__ == "__main__":
    main()
