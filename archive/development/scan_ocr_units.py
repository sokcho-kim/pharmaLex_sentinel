# -*- coding: utf-8 -*-
"""
PDF OCR 후 단위/기호 깨짐 전수 스캔 → CSV로 내보내기
- 요구 라이브러리: PyMuPDF (fitz), pandas
    pip install pymupdf pandas
"""

import re
import csv
from pathlib import Path
import fitz  # PyMuPDF
import pandas as pd

# ====== 설정 ======
PDF_PATH = r"C:\Jimin\pharmaLex_sentinel\data\요양급여의 적용기준 및방법에 관한 세부사항(약제).pdf"  # 대상 PDF 경로
OUT_CSV  = "./out/ocr_unit_anomalies_scan.csv"

# 작은 g(그램)을 ㎍(마이크로그램) 오인으로 의심할 기준값 (너무 큰 g는 진짜 g일 가능성 높음)
GRAM_SUSPECT_THRESHOLD = 100  # 100g 이하이면 의심(도메인에 맞게 조정)

# g가 ㎍일 확률을 더 올려주는 주변 '제형/맥락' 한국어 키워드(선택)
FORM_HINTS = [
    "이식제","정","캡슐","현탁","시럽","액","주","흡입","분무","패치","장용","서방","안연고","점안","현탁액",
    "흡입제","스프레이","연고","겔","로션","시럽제","과립","산제","분말","점비","좌제","점이","주사","주사용"
]

# ====== 패턴 정의 ======
# 수치+단위류
PATTERNS = {
    "micro_ascii": re.compile(r"\b(\d+(?:\.\d+)?)\s*(?:ug|mcg)\b", re.IGNORECASE),
    "micro_symbol": re.compile(r"\b(\d+(?:\.\d+)?)\s*㎍\b"),
    "milli_ascii": re.compile(r"\b(\d+(?:\.\d+)?)\s*mg\b", re.IGNORECASE),
    "milli_symbol": re.compile(r"\b(\d+(?:\.\d+)?)\s*㎎\b"),
    "gram_ascii":  re.compile(r"\b(\d+(?:\.\d+)?)\s*g\b", re.IGNORECASE),
    "ml_ascii":    re.compile(r"\b(\d+(?:\.\d+)?)\s*ml\b", re.IGNORECASE),
    "ml_symbol":   re.compile(r"\b(\d+(?:\.\d+)?)\s*㎖\b"),
    "iu_ascii":    re.compile(r"\b(\d+(?:\.\d+)?)\s*iu\b", re.IGNORECASE),
    # 그리스 문자 (정상 검출 포함)
    "greek_letters": re.compile(r"[αβγμ]"),
}

# α/β/γ가 a/b/g로 깨졌을 가능성 (보수적 규칙: 영문자 단독 하이픈 접두 등)
GREEK_MIS_OCR = {
    "alpha_like": re.compile(r"\b(a-|\balpha\b)", re.IGNORECASE),
    "beta_like":  re.compile(r"\b(b-|\bbeta\b)", re.IGNORECASE),
    "gamma_like": re.compile(r"\b(g-|\bgamma\b)", re.IGNORECASE),
    # μ 단독: 문맥상 단위와 붙어야 하는데 빠진 케이스(예: μ g / μ l 등에서 분리되거나, 그냥 μ만 남음)
    "mu_alone":   re.compile(r"\bμ\b"),
}

def get_context(text: str, start: int, end: int, window: int = 60) -> str:
    s = max(0, start - window)
    e = min(len(text), end + window)
    snippet = text[s:e].replace("\n", " ")
    return re.sub(r"\s+", " ", snippet).strip()

def has_form_hint(context: str) -> bool:
    return any(h in context for h in FORM_HINTS)

def classify_and_suggest(kind: str, value: str, context: str):
    """
    kind: 패턴 키
    value: 수치/문자 추출 값(숫자 문자열 or 기호)
    context: 주변 문맥
    return: (classification, suggested_fix, reason)
    """
    # 기본값
    classification = "info"
    suggested = ""
    reason = ""

    if kind == "micro_ascii":
        classification = "normalize"
        suggested = f"{value} ㎍"
        reason = "ASCII 'ug/mcg' → 기호 '㎍'로 정규화 권장."

    elif kind == "micro_symbol":
        classification = "ok"
        suggested = f"{value} ㎍"
        reason = "정상 마이크로그램 표기."

    elif kind == "milli_ascii":
        classification = "ok"
        suggested = f"{value} mg"
        reason = "정상 밀리그램 표기(ASCII)."

    elif kind == "milli_symbol":
        classification = "ok"
        suggested = f"{value} ㎎"
        reason = "정상 밀리그램 표기(기호)."

    elif kind == "gram_ascii":
        # g를 ㎍ 오인으로 의심: 값이 비현실적으로 작거나, 제형 힌트가 있으면 의심도 상승
        try:
            num = float(value)
        except:
            num = None

        if num is not None and num <= GRAM_SUSPECT_THRESHOLD and has_form_hint(context):
            classification = "suspect_micro_as_g"
            suggested = f"{value} ㎍"
            reason = f"g(그램)로 OCR되었으나 제형 힌트+소량({value}g) → ㎍(마이크로그램) 오인 가능."
        elif num is not None and num <= GRAM_SUSPECT_THRESHOLD:
            classification = "review_micro_as_g"
            suggested = f"{value} ㎍ (검토)"
            reason = f"g(그램)로 OCR되었으나 소량({value}g) → ㎍ 오인 가능성. 문맥 검토 요망."
        else:
            classification = "ok_or_large_g"
            suggested = f"{value} g"
            reason = "값이 커서 실제 g(그램)일 가능성이 높음."

    elif kind in ("ml_ascii", "ml_symbol"):
        classification = "ok"
        suggested = f"{value} ㎖" if kind == "ml_symbol" else f"{value} mL"
        reason = "정상 밀리리터 표기."

    elif kind == "iu_ascii":
        classification = "ok"
        suggested = f"{value} IU"
        reason = "정상 국제단위 표기."

    elif kind == "greek_letters":
        classification = "ok"
        suggested = "(그대로)"
        reason = "그리스 문자 정상 검출."

    elif kind in ("alpha_like", "beta_like", "gamma_like"):
        classification = "suspect_greek_broken"
        letter = {"alpha_like": "α", "beta_like": "β", "gamma_like": "γ"}[kind]
        suggested = f"(문맥상 {letter} 검토)"
        reason = f"a/b/g 형태가 {letter}로 깨졌을 가능성."

    elif kind == "mu_alone":
        classification = "suspect_mu_alone"
        suggested = "(문맥상 ㎍ 또는 ㎖ 검토)"
        reason = "μ 단독 검출 → 단위 기호 깨짐 가능성."

    else:
        classification = "info"
        suggested = ""
        reason = "기타 정보."

    return classification, suggested, reason

def scan_pdf(pdf_path: str):
    doc = fitz.open(pdf_path)
    rows = []

    for pno in range(len(doc)):
        page = doc[pno]
        text = page.get_text("text")

        # 1) 수치+단위 패턴 모두 스캔
        for key, pat in PATTERNS.items():
            for m in pat.finditer(text):
                if key in ("micro_ascii","micro_symbol","milli_ascii","milli_symbol","gram_ascii","ml_ascii","ml_symbol","iu_ascii"):
                    num = m.group(1)
                    match_txt = m.group(0)
                else:
                    num = ""
                    match_txt = m.group(0)

                ctx = get_context(text, m.start(), m.end())
                cls, sug, rsn = classify_and_suggest(key, num if num else match_txt, ctx)

                rows.append({
                    "page": pno + 1,
                    "match": match_txt,
                    "classification": cls,
                    "suggested_fix": sug,
                    "reason": rsn,
                    "context": ctx
                })

        # 2) 그리스 문자 오인 의심(a-, b-, g-, alpha/beta/gamma)
        for key, pat in GREEK_MIS_OCR.items():
            for m in pat.finditer(text):
                match_txt = m.group(0)
                ctx = get_context(text, m.start(), m.end())
                cls, sug, rsn = classify_and_suggest(key, match_txt, ctx)
                rows.append({
                    "page": pno + 1,
                    "match": match_txt,
                    "classification": cls,
                    "suggested_fix": sug,
                    "reason": rsn,
                    "context": ctx
                })

    return rows

def main():
    rows = scan_pdf(PDF_PATH)
    df = pd.DataFrame(rows).sort_values(["classification","page"]).reset_index(drop=True)

    # 우선 ‘의심’ 위주로 위로 정렬되게 가중 정렬(선택)
    priority = {
        "suspect_micro_as_g": 1,
        "review_micro_as_g": 2,
        "suspect_greek_broken": 3,
        "suspect_mu_alone": 4,
        "normalize": 5,
        "ok": 6,
        "ok_or_large_g": 7,
        "info": 9,
    }
    df["prio"] = df["classification"].map(priority).fillna(99).astype(int)
    df = df.sort_values(["prio","page"]).drop(columns=["prio"])

    df.to_csv(OUT_CSV, index=False, encoding="utf-8-sig")
    print(f"[완료] CSV 저장: {OUT_CSV} (총 {len(df)}건)")

if __name__ == "__main__":
    main()
