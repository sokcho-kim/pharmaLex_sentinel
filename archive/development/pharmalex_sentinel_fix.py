# -*- coding: utf-8 -*-
"""
PharmaLex Sentinel — 단위/기호 정규화 · 패치 파이프라인 (최종본)
- 목적:
  1) data/요양심사약제_후처리.xlsx 전수 검사 (오류 개수/위치 파악)
  2) out/ocr_unit_anomalies_scan.csv 활용(패턴/맥락) + 룰 기반 검증
  3) 오류 자동 패치본 저장 + 어디를 어떻게 고쳤는지 CSV 로그 생성

- 권장 실행 위치: C:\Jimin\pharmaLex_sentinel\
- 입력:
    data/요양심사약제_후처리.xlsx
    out/ocr_unit_anomalies_scan.csv
    (선택) data/요양급여의 적용기준 및방법에 관한 세부사항(약제).pdf  # 이미 스캔 반영됨
- 출력:
    out/요양심사약제_후처리_수정본.xlsx
    out/error_corrections.csv
    out/summary_report.md
"""

import os
import re
import json
import math
import datetime as dt
import pandas as pd

# ===================== 사용자 설정 =====================
BASE_DIR = r"C:\Jimin\pharmaLex_sentinel"  # 형님 환경 경로
IN_EXCEL = os.path.join(BASE_DIR, r"data\요양심사약제_후처리.xlsx")
OCR_CSV  = os.path.join(BASE_DIR, r"out\ocr_unit_anomalies_scan.csv")

OUT_DIR  = os.path.join(BASE_DIR, "out")
OUT_EXCEL = os.path.join(OUT_DIR, r"요양심사약제_후처리_수정본.xlsx")
OUT_LOG   = os.path.join(OUT_DIR, "error_corrections.csv")
OUT_SUMMARY = os.path.join(OUT_DIR, "summary_report.md")

# g→㎍ 의심값 상한(도메인 조정 가능)
GRAM_SUSPECT_THRESHOLD = 100.0

# 제형 힌트(있으면 g→㎍ 교정 신뢰도↑)
FORM_KEYWORDS = [
    "정","주","주사","시럽","이식제","캡슐","패치","외용제","점안액","연고",
    "겔","로션","현탁","현탁액","흡입","분무","스프레이","장용","서방","좌제","과립","산제","분말","점비"
]

# g 단위를 실험실 수치/검사값 맥락으로 판단하는 부정 키워드/패턴(있으면 교정 금지)
LAB_NEG_PATTERNS = [
    r"g\/dl", r"g\/l", r"g\/24h", r"g\/day", r"g\/g", r"g\/m2", r"g\/m²",
    r"\bhb\b", r"\bhct\b", "헤모글로빈", "혈장", "단백뇨", "경구당부하"
]

# ===================== 유틸 함수 =====================

def safe_lower(s: str) -> str:
    return s.lower() if isinstance(s, str) else s

def contains_any(text: str, keywords) -> bool:
    t = text if isinstance(text, str) else str(text)
    return any(k in t for k in keywords)

def regex_any(text: str, patterns) -> bool:
    t = text if isinstance(text, str) else str(text)
    return any(re.search(pat, t, flags=re.IGNORECASE) for pat in patterns)

def load_ocr_anomalies(csv_path: str) -> pd.DataFrame:
    if not os.path.exists(csv_path):
        print(f"[경고] OCR 스캔 CSV가 없습니다: {csv_path}  (룰 기반만 동작)")
        return pd.DataFrame()
    df = pd.read_csv(csv_path)
    # 필요한 칼럼 보정
    for col in ["match","classification","suggested_fix","context","page","reason"]:
        if col not in df.columns:
            df[col] = ""
    return df

def build_ascii_micro_regex():
    # 예: "20mcg", "5 ug", "0.7UG" -> "숫자 + (mcg|ug)"
    return re.compile(r"\b(\d+(?:\.\d+)?)\s*(mcg|ug)\b", flags=re.IGNORECASE)

ASCII_MICRO_RE = build_ascii_micro_regex()
G_VALUE_RE = re.compile(r"(\d+(?:\.\d+)?)\s*g(?![a-zA-Z])")  # 숫자 + g(뒤에 영문 없음)

def normalize_ascii_micro(cell_text: str):
    """
    ASCII ug/mcg → 기호 ㎍ 로 치환. (안전)
    반환: (변경여부, 변경후문자열, [로그리스트])
    """
    logs = []
    def repl(m):
        val = m.group(1)
        unit = m.group(2)
        before = m.group(0)
        after = f"{val} ㎍"
        logs.append({
            "rule": "ascii_micro",
            "before": before,
            "after": after,
            "detail": f"{unit} -> ㎍"
        })
        return after

    new_text, n = ASCII_MICRO_RE.subn(repl, cell_text)
    return (n > 0, new_text, logs)

def should_convert_g_to_micro(cell_text: str, g_value: float) -> (bool, str):
    """
    g → ㎍ 치환을 해도 되는지 판단.
    - 값이 작음(<=GRAM_SUSPECT_THRESHOLD)
    - 제형 키워드가 있음
    - 검사/실험실 맥락 부정 패턴이 없음
    """
    s = safe_lower(cell_text)
    has_form = contains_any(s, [k for k in FORM_KEYWORDS])  # 한글 키워드는 소문자 처리 영향 없음
    has_lab = regex_any(s, LAB_NEG_PATTERNS)
    if g_value <= GRAM_SUSPECT_THRESHOLD and has_form and not has_lab:
        return True, "value<=threshold & has_form & no_lab"
    return False, f"value={g_value}, form={has_form}, lab={has_lab}"

def normalize_g_to_micro(cell_text: str):
    """
    '숫자 g' → 조건부 '숫자 ㎍'
    반환: (변경여부, 변경후문자열, [로그리스트], [검토필요로그])
    """
    logs, reviews = [], []
    changed = False

    def repl(m):
        nonlocal changed
        val = m.group(1)
        before = m.group(0)           # 예: "3g"
        ok, reason = should_convert_g_to_micro(cell_text, float(val))
        if ok:
            after = f"{val} ㎍"
            logs.append({
                "rule": "g_to_micro_conditional",
                "before": before,
                "after": after,
                "detail": reason
            })
            changed = True
            return after
        else:
            # 자동 미적용, 검토 목록에 올림
            reviews.append({
                "rule": "g_to_micro_review",
                "before": before,
                "suggested": f"{val} ㎍ (검토)",
                "detail": reason
            })
            return before  # 교정하지 않음

    new_text = G_VALUE_RE.sub(repl, cell_text)
    return changed, new_text, logs, reviews

# ===================== 핵심 처리 =====================

def process_dataframe(df: pd.DataFrame, sheet_name: str, ocr_df: pd.DataFrame):
    """
    각 시트 DF에 대해 텍스트 컬럼 전수 검사 → 자동교정/검토 목록/로그 생성
    """
    df_out = df.copy()
    corrections = []   # 자동 교정 로그
    reviews = []       # 검토 필요 로그
    total_cells = 0
    changed_cells = 0

    text_cols = df_out.select_dtypes(include=["object"]).columns.tolist()

    # OCR CSV가 제공하는 'match'가 있다면 참고(통계용). 교정은 룰 기반으로만.
    known_suspicious_strings = set(s for s in ocr_df["match"].astype(str).unique()) if not ocr_df.empty else set()

    for col in text_cols:
        for idx, val in df_out[col].items():
            if pd.isna(val):
                continue
            cell = str(val)
            orig = cell
            total_cells += 1

            cell_logs = []
            cell_reviews = []

            # 1) ASCII micro 교정
            asc_changed, cell, asc_logs = normalize_ascii_micro(cell)
            if asc_changed:
                cell_logs.extend(asc_logs)

            # 2) g → ㎍ 조건부
            g_changed, cell, g_logs, g_reviews = normalize_g_to_micro(cell)
            if g_changed:
                cell_logs.extend(g_logs)
            if g_reviews:
                cell_reviews.extend(g_reviews)

            # 변경 반영 / 로그 적재
            if cell != orig:
                changed_cells += 1
                df_out.at[idx, col] = cell
                for lg in cell_logs:
                    corrections.append({
                        "sheet": sheet_name,
                        "row_idx": idx,
                        "column": col,
                        "rule": lg["rule"],
                        "before": lg["before"],
                        "after": lg["after"],
                        "detail": lg["detail"],
                        "had_ocr_match": any(lg["before"] == s for s in known_suspicious_strings)
                    })

            # 검토만 필요한 경우도 기록
            for rv in cell_reviews:
                reviews.append({
                    "sheet": sheet_name,
                    "row_idx": idx,
                    "column": col,
                    "rule": rv["rule"],
                    "before": rv["before"],
                    "suggested": rv["suggested"],
                    "detail": rv["detail"],
                    "cell_excerpt": orig[:120]
                })

    return df_out, corrections, reviews, total_cells, changed_cells

def process_workbook(in_excel: str, ocr_csv: str):
    if not os.path.exists(in_excel):
        raise FileNotFoundError(f"입력 엑셀 없음: {in_excel}")

    os.makedirs(OUT_DIR, exist_ok=True)

    # OCR 스캔 결과 로드 (통계/참고)
    ocr_df = load_ocr_anomalies(ocr_csv)

    xls = pd.ExcelFile(in_excel)
    all_sheets = xls.sheet_names

    writer = pd.ExcelWriter(OUT_EXCEL, engine="openpyxl")
    all_corrections = []
    all_reviews = []
    grand_total = 0
    grand_changed = 0

    for s in all_sheets:
        df = pd.read_excel(in_excel, sheet_name=s, dtype=str)  # 모든 컬럼 문자열로(치환 안정성↑)
        df_fixed, corr, rvw, tot, chg = process_dataframe(df, s, ocr_df)
        df_fixed.to_excel(writer, sheet_name=s, index=False)

        all_corrections.extend(corr)
        all_reviews.extend(rvw)
        grand_total += tot
        grand_changed += chg

    writer.close()

    # 로그 저장
    log_df = pd.DataFrame(all_corrections)
    review_df = pd.DataFrame(all_reviews)

    if not log_df.empty:
        log_df.to_csv(OUT_LOG, index=False, encoding="utf-8-sig")

    # 요약 리포트
    ts = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    auto_cnt = len(log_df) if not log_df.empty else 0
    review_cnt = len(review_df) if not review_df.empty else 0

    summary = []
    summary.append(f"# PharmaLex Sentinel 정규화 리포트")
    summary.append(f"- 실행시각: {ts}")
    summary.append(f"- 입력 엑셀: `{IN_EXCEL}`")
    summary.append(f"- OCR 스캔: `{OCR_CSV}`")
    summary.append("")
    summary.append(f"## 처리 요약")
    summary.append(f"- 전체 검사 셀 수: **{grand_total}**")
    summary.append(f"- 변경된 셀 수: **{grand_changed}**")
    summary.append(f"- 자동 교정 로그 수(건별): **{auto_cnt}**")
    summary.append(f"- 사람 검토 필요 수(건별): **{review_cnt}**")
    summary.append("")
    summary.append(f"## 산출물")
    summary.append(f"- 교정본 엑셀: `{OUT_EXCEL}`")
    summary.append(f"- 교정 로그 CSV: `{OUT_LOG}`" if auto_cnt else "- 교정 로그 CSV: (변경 없음)")
    summary.append(f"- 검토 목록: 아래 표 (샘플 50건)")
    summary.append("")
    if not review_df.empty:
        summary.append(review_df.head(50).to_markdown(index=False))
    else:
        summary.append("_검토 필요 없음_")

    with open(OUT_SUMMARY, "w", encoding="utf-8") as f:
        f.write("\n".join(summary))

    return {
        "sheets": all_sheets,
        "total_cells": grand_total,
        "changed_cells": grand_changed,
        "auto_logs": auto_cnt,
        "review_logs": review_cnt,
        "out_excel": OUT_EXCEL,
        "out_log": OUT_LOG,
        "out_summary": OUT_SUMMARY
    }

# ===================== 실행부 =====================

if __name__ == "__main__":
    info = process_workbook(IN_EXCEL, OCR_CSV)
    print(json.dumps(info, ensure_ascii=False, indent=2))
