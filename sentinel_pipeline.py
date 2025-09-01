# -*- coding: utf-8 -*-
"""
PharmaLex Sentinel: 원본→(Step1 클린)→(Step2 정규화) 파이프라인

입력:
  data/요양심사약제_후처리.xlsx            (원본)
  out/ocr_unit_anomalies_scan.csv          (OCR 단위 이상 스캔 결과)
  (선택) data/mapping.csv                  (� 등 깨진 문자 수동 매핑표)

출력:
  out/invalid_char_report.csv              (� 전수 리포트)
  out/요양심사약제_후처리_clean.xlsx       (1단계 클린본; mapping 있으면 생성)
  out/요양심사약제_후처리_normalized.xlsx  (2단계 최종본)
  out/error_corrections.csv                (치환 로그)
  out/summary_report.md                    (요약 리포트)
"""
import os, re, json, datetime as dt
import pandas as pd

# ---------------- 경로 설정 ----------------
BASE = r"C:\Jimin\pharmaLex_sentinel"
IN_XLSX = os.path.join(BASE, r"data\요양심사약제_후처리.xlsx")         # 원본
OCR_CSV = os.path.join(BASE, r"out\ocr_unit_anomalies_scan.csv")     # 기존 생성본
MAPPING_CSV = os.path.join(BASE, r"data\mapping.csv")                # 선택(수동 매핑표)

OUT_DIR = os.path.join(BASE, "out")
REPORT_INVALID = os.path.join(OUT_DIR, "invalid_char_report.csv")
CLEAN_XLSX = os.path.join(OUT_DIR, "요양심사약제_후처리_clean.xlsx")
NORM_XLSX  = os.path.join(OUT_DIR, "요양심사약제_후처리_normalized.xlsx")
LOG_CSV    = os.path.join(OUT_DIR, "error_corrections.csv")
SUMMARY_MD = os.path.join(OUT_DIR, "summary_report.md")

# ---------------- Step1: � 탐지/치환 ----------------
def scan_invalid_chars(excel_path: str, out_report: str) -> int:
    xls = pd.ExcelFile(excel_path)
    rows = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(excel_path, sheet_name=sheet, dtype=str)
        for col in df.columns:
            for idx, val in df[col].items():
                if pd.isna(val): 
                    continue
                val = str(val)
                if "�" in val:
                    rows.append({
                        "sheet": sheet,
                        "row": idx + 2,   # 헤더 감안
                        "column": col,
                        "value": val,
                        "count_in_cell": val.count("�")
                    })
    rep = pd.DataFrame(rows)
    os.makedirs(os.path.dirname(out_report), exist_ok=True)
    rep.to_csv(out_report, index=False, encoding="utf-8-sig")
    return len(rep)

def load_mapping(mapping_csv: str):
    """
    mapping.csv 예시 (UTF-8-SIG 저장):
    before,after,notes
    �,㎍,주요 단위 복구
    """
    if not os.path.exists(mapping_csv):
        return []
    m = pd.read_csv(mapping_csv, dtype=str).fillna("")
    pairs = []
    for _, r in m.iterrows():
        pairs.append((str(r["before"]), str(r["after"])))
    return pairs

def apply_mapping_to_workbook(in_xlsx: str, out_xlsx: str, pairs: list):
    xls = pd.ExcelFile(in_xlsx)
    writer = pd.ExcelWriter(out_xlsx, engine="openpyxl")
    for sheet in xls.sheet_names:
        df = pd.read_excel(in_xlsx, sheet_name=sheet, dtype=str)
        # 문자형 컬럼에만 치환
        for col in df.columns:
            if df[col].dtype == object:
                s = df[col].astype(str)
                for before, after in pairs:
                    s = s.str.replace(before, after, regex=False)
                df[col] = s
        df.to_excel(writer, sheet_name=sheet, index=False)
    writer.close()

# ---------------- Step2: 정규화(단위/기호) ----------------
GRAM_SUSPECT_THRESHOLD = 100.0
FORM_KEYWORDS = [
    "정","주","주사","시럽","이식제","캡슐","패치","외용제","점안액","연고",
    "겔","로션","현탁","현탁액","흡입","분무","스프레이","장용","서방","좌제","과립","산제","분말","점비"
]
LAB_NEG_PATTERNS = [
    r"g\/dl", r"g\/l", r"g\/24h", r"g\/day", r"g\/g", r"g\/m2", r"g\/m²",
    r"\bhb\b", r"\bhct\b", "헤모글로빈", "혈장", "단백뇨", "경구당부하"
]
ASCII_MICRO_RE = re.compile(r"\b(\d+(?:\.\d+)?)\s*(mcg|ug)\b", re.IGNORECASE)
G_VALUE_RE = re.compile(r"(\d+(?:\.\d+)?)\s*g(?![a-zA-Z])")

def contains_any(text: str, keywords) -> bool:
    t = text if isinstance(text, str) else str(text)
    return any(k in t for k in keywords)

def regex_any(text: str, patterns) -> bool:
    t = text if isinstance(text, str) else str(text)
    return any(re.search(p, t, flags=re.IGNORECASE) for p in patterns)

def normalize_ascii_micro(cell_text: str):
    logs = []
    def repl(m):
        val = m.group(1); unit = m.group(2)
        before = m.group(0); after = f"{val} ㎍"
        logs.append(("ascii_micro", before, after, f"{unit} -> ㎍"))
        return after
    new_text, n = ASCII_MICRO_RE.subn(repl, cell_text)
    return (n > 0, new_text, logs)

def should_convert_g_to_micro(context_text: str, g_value: float):
    has_form = contains_any(context_text, FORM_KEYWORDS)
    has_lab  = regex_any(context_text, LAB_NEG_PATTERNS)
    if g_value <= GRAM_SUSPECT_THRESHOLD and has_form and not has_lab:
        return True, "value<=threshold & has_form & no_lab"
    return False, f"value={g_value}, form={has_form}, lab={has_lab}"

def normalize_g_to_micro(cell_text: str):
    logs, reviews = [], []
    changed = False
    def repl(m):
        nonlocal changed
        val = m.group(1); before = m.group(0)
        ok, reason = should_convert_g_to_micro(cell_text, float(val))
        if ok:
            after = f"{val} ㎍"
            logs.append(("g_to_micro_conditional", before, after, reason))
            changed = True
            return after
        else:
            reviews.append(("g_to_micro_review", before, f"{val} ㎍ (검토)", reason))
            return before
    new_text = G_VALUE_RE.sub(repl, cell_text)
    return changed, new_text, logs, reviews

def normalize_workbook(in_xlsx: str, ocr_csv: str, out_xlsx: str, out_log_csv: str, out_summary_md: str):
    os.makedirs(os.path.dirname(out_xlsx), exist_ok=True)
    xls = pd.ExcelFile(in_xlsx)
    writer = pd.ExcelWriter(out_xlsx, engine="openpyxl")
    all_logs, all_reviews = [], []
    total_cells = changed_cells = 0

    for sheet in xls.sheet_names:
        df = pd.read_excel(in_xlsx, sheet_name=sheet, dtype=str)
        df_out = df.copy()
        for col in df.columns:
            for idx, val in df[col].items():
                if pd.isna(val): 
                    continue
                cell = str(val); orig = cell
                total_cells += 1

                # 1) ASCII ug/mcg
                asc_changed, cell, asc_logs = normalize_ascii_micro(cell)
                for rule, before, after, detail in asc_logs:
                    all_logs.append({
                        "sheet": sheet, "row_idx": idx, "column": col,
                        "rule": rule, "before": before, "after": after, "detail": detail
                    })

                # 2) g → ㎍ 조건부
                g_changed, cell, g_logs, g_reviews = normalize_g_to_micro(cell)
                for rule, before, after, detail in g_logs:
                    all_logs.append({
                        "sheet": sheet, "row_idx": idx, "column": col,
                        "rule": rule, "before": before, "after": after, "detail": detail
                    })
                for rule, before, suggested, detail in g_reviews:
                    all_reviews.append({
                        "sheet": sheet, "row_idx": idx, "column": col,
                        "rule": rule, "before": before, "suggested": suggested,
                        "detail": detail, "cell_excerpt": orig[:120]
                    })

                if cell != orig:
                    changed_cells += 1
                    df_out.at[idx, col] = cell

        df_out.to_excel(writer, sheet_name=sheet, index=False)
    writer.close()

    pd.DataFrame(all_logs).to_csv(out_log_csv, index=False, encoding="utf-8-sig")

    ts = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(out_summary_md, "w", encoding="utf-8") as f:
        f.write(f"# PharmaLex Sentinel 정규화 리포트\n- 실행시각: {ts}\n")
        f.write(f"- 입력 엑셀: `{in_xlsx}`\n- OCR 스캔: `{ocr_csv}`\n\n")
        f.write("## 처리 요약\n")
        f.write(f"- 전체 검사 셀 수: **{total_cells}**\n")
        f.write(f"- 변경된 셀 수: **{changed_cells}**\n")
        f.write(f"- 자동 교정 로그 수: **{len(all_logs)}**\n")
        f.write(f"- 사람 검토 필요 수: **{len(all_reviews)}**\n\n")
        if all_reviews:
            pd.DataFrame(all_reviews).head(50).to_csv(
                os.path.join(os.path.dirname(out_summary_md), "review_samples.csv"),
                index=False, encoding="utf-8-sig"
            )
            f.write("- 검토 샘플: `review_samples.csv` 참조\n")
        else:
            f.write("- 검토 필요 없음\n")

def main():
    os.makedirs(OUT_DIR, exist_ok=True)

    # Step 1: invalid scan
    cnt = scan_invalid_chars(IN_XLSX, REPORT_INVALID)
    print(f"[Step1] invalid_char_report.csv 생성 (rows={cnt}) -> {REPORT_INVALID}")

    # Step 1b: optional mapping
    pairs = load_mapping(MAPPING_CSV)
    if pairs:
        apply_mapping_to_workbook(IN_XLSX, CLEAN_XLSX, pairs)
        print(f"[Step1] mapping.csv 적용 -> {CLEAN_XLSX}")
        step2_input = CLEAN_XLSX
    else:
        print("[Step1] mapping.csv 없음 → 원본으로 Step2 진행")
        step2_input = IN_XLSX

    # Step 2: normalize
    normalize_workbook(step2_input, OCR_CSV, NORM_XLSX, LOG_CSV, SUMMARY_MD)
    print(f"[Step2] 정규화 완료 -> {NORM_XLSX}")
    print(f"[LOG] {LOG_CSV}")
    print(f"[SUMMARY] {SUMMARY_MD}")

if __name__ == "__main__":
    main()
