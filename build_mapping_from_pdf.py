# -*- coding: utf-8 -*-
"""
PDF 역검색으로 �(U+FFFD) 교정 후보 제안 → mapping.csv 반자동 생성

입력:
  data/요양심사약제_후처리.xlsx        # 원본(또는 clean 이전본)
  data/요양급여의 적용기준 및방법에 관한 세부사항(약제).pdf
출력:
  out/mapping_candidates.csv             # 후보/근거(페이지) 제안표
  data/mapping.csv                       # (선택) 확정본 생성용; 아래 '확정 단계' 참고
"""

import re, os, csv
from collections import defaultdict, Counter
import pandas as pd
import fitz  # PyMuPDF

# 경로
BASE = r"C:\Jimin\pharmaLex_sentinel"
IN_XLSX = os.path.join(BASE, r"data\요양심사약제_후처리.xlsx")
IN_PDF  = os.path.join(BASE, r"data\요양급여의 적용기준 및방법에 관한 세부사항(약제).pdf")
OUT_DIR = os.path.join(BASE, "out")
CAND_CSV = os.path.join(OUT_DIR, "mapping_candidates.csv")
FINAL_MAP = os.path.join(BASE, r"data\mapping.csv")

# � 대체 후보(필요 시 추가)
CANDIDATES = ["㎍","㎎","㎖","α","β","γ","μ","-","·","×","~","/"]

# 검색 옵션
CONTEXT_CHARS = 12        # � 좌우로 붙일 문맥 길이
CASE_INSENSITIVE = True   # 대소문자 무시

def load_pdf_text_by_page(pdf_path):
    doc = fitz.open(pdf_path)
    pages = []
    for p in doc:
        txt = p.get_text("text")
        if CASE_INSENSITIVE: txt = txt.lower()
        # 공백 정규화
        txt = re.sub(r"\s+", " ", txt)
        pages.append(txt)
    return pages

def normalize(s):
    if not isinstance(s, str): s = str(s)
    s = s.replace("\n"," ")
    s = re.sub(r"\s+"," ", s)
    return s.lower() if CASE_INSENSITIVE else s

def iter_fffd_cells(xlsx_path):
    xls = pd.ExcelFile(xlsx_path)
    for sheet in xls.sheet_names:
        df = pd.read_excel(xlsx_path, sheet_name=sheet, dtype=str)
        for col in df.columns:
            for ridx, val in df[col].items():
                if pd.isna(val): continue
                if "�" in str(val):
                    yield sheet, col, ridx, str(val)

def build_regex_from_context(text_with_fffd, pos, candidate):
    """
    text_with_fffd에서 pos 위치의 � 하나를 candidate로 치환한 '느슨한' 정규식 패턴 생성
    - 좌우로 CONTEXT_CHARS 만큼 문맥을 사용 (공백은 \s+로 느슨하게)
    """
    left = text_with_fffd[max(0, pos - CONTEXT_CHARS):pos]
    right = text_with_fffd[pos+1: pos+1+CONTEXT_CHARS]

    # 정규식 이스케이프 + 공백 느슨화
    def esc_relax(s):
        s = re.escape(s)
        s = s.replace(r"\ ", r"\s+")
        return s

    pat = esc_relax(left) + re.escape(candidate) + esc_relax(right)
    return re.compile(pat)

def scan_candidates_in_pdf(pages_text, text_val):
    """
    셀 문자열(text_val) 안의 모든 �에 대해 후보별로 PDF 페이지에서 매칭 수를 센다.
    반환: dict(candidate -> list of (page_idx, count)) 와 최고의 후보 집계
    """
    t = normalize(text_val)
    # � 위치들
    pos_list = [m.start() for m in re.finditer("�", t)]
    if not pos_list:
        return {}

    page_hits = {cand: Counter() for cand in CANDIDATES}

    for pos in pos_list:
        for cand in CANDIDATES:
            rgx = build_regex_from_context(t, pos, cand)
            for pidx, page_txt in enumerate(pages_text):
                # 페이지에서 패턴 매칭 수
                hits = len(list(rgx.finditer(page_txt)))
                if hits > 0:
                    page_hits[cand][pidx+1] += hits  # 1-based page

    # 후보 요약 (페이지/카운트)
    summary = {}
    for cand, ctr in page_hits.items():
        total = sum(ctr.values())
        if total > 0:
            # 상위 3개 페이지만 요약
            top3 = ", ".join([f"p{p}×{c}" for p,c in ctr.most_common(3)])
            summary[cand] = {"total": total, "top_pages": top3}
    return summary

def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    print("[1/3] PDF 로딩…")
    pages = load_pdf_text_by_page(IN_PDF)

    rows = []
    print("[2/3] 엑셀 내 � 셀 스캔…")
    for sheet, col, ridx, val in iter_fffd_cells(IN_XLSX):
        cand_stats = scan_candidates_in_pdf(pages, val)
        # 후보가 하나도 안 잡히면 공란으로
        if not cand_stats:
            rows.append({
                "sheet": sheet, "row": ridx+2, "column": col,
                "value": val, "best_candidate": "",
                "candidate_scores": "",
                "final_after": ""  # <- 형님이 여기 채우면 mapping.csv 생성 가능
            })
            continue

        # 총합 빈도 최댓값 후보 선택
        best = max(cand_stats.items(), key=lambda kv: kv[1]["total"])[0]
        scores = " | ".join([f"{c}:{d['total']}({d['top_pages']})" for c, d in sorted(cand_stats.items(), key=lambda kv: -kv[1]["total"])])

        rows.append({
            "sheet": sheet,
            "row": ridx+2,  # 엑셀 행 번호 보정
            "column": col,
            "value": val,
            "best_candidate": best,
            "candidate_scores": scores,
            "final_after": ""  # 사람이 최종 확정
        })

    print("[3/3] 후보표 저장…")
    pd.DataFrame(rows).to_csv(CAND_CSV, index=False, encoding="utf-8-sig")
    print(f"→ {CAND_CSV}")
    print("\n이제 아래 '확정 단계'를 따라 주세요.")

if __name__ == "__main__":
    main()
