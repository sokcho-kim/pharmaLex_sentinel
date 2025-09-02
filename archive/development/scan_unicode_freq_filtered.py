# -*- coding: utf-8 -*-
import os, pandas as pd
from collections import Counter, defaultdict

BASE = r"C:\Jimin\pharmaLex_sentinel"
IN_XLSX = os.path.join(BASE, r"data\요양심사약제_후처리.xlsx")
OUT_DIR = os.path.join(BASE, "out_scan"); os.makedirs(OUT_DIR, exist_ok=True)

FREQ_CSV    = os.path.join(OUT_DIR, "unicode_freq_filtered.csv")
SAMPLES_CSV = os.path.join(OUT_DIR, "unicode_samples_filtered.csv")

# 1) 반드시 보고 싶은 의심 문자(allowlist)
TARGET_CHARS = set([
    "�", "㎍", "㎎", "㎖", "α", "β", "γ", "μ",
    "°", "±", "≤", "≥", "·", "×", "–", "—", "™", "®"
])

# 2) 한글/숫자/영문/일반문장부호 등은 무시(denylist 범위)
def is_korean(cp):  # 한글 범위 제외
    return (0xAC00 <= cp <= 0xD7A3) or (0x1100 <= cp <= 0x11FF) or (0x3130 <= cp <= 0x318F)

def is_basic_ascii(cp):  # 영문/숫자/기본문장부호
    return 0x20 <= cp <= 0x7E

# 3) 우리가 실제로 카운트할지 여부
def is_suspicious_char(ch):
    cp = ord(ch)
    if ch in TARGET_CHARS:             # 우선 타깃이면 무조건 포함
        return True
    if is_korean(cp) or is_basic_ascii(cp):
        return False                    # 한글/기본ASCII는 제외
    # 그 밖의 비한글/비ASCII 기호는 의심 후보로 포함 (예: 희귀 특수문자)
    return True

def scan(path):
    from collections import Counter, defaultdict
    per_char = Counter()
    samples = defaultdict(list)
    xls = pd.ExcelFile(path)

    for sheet in xls.sheet_names:  # 시트 한 개여도 일반화
        df = pd.read_excel(path, sheet_name=sheet, dtype=str)
        for r in range(len(df)):
            for col in df.columns:
                v = df.iat[r, df.columns.get_loc(col)]
                if pd.isna(v): continue
                s = str(v)
                for ch in s:
                    if is_suspicious_char(ch):
                        cp = ord(ch)
                        per_char[cp] += 1
                        if len(samples[cp]) < 5:
                            samples[cp].append({
                                "sheet": sheet, "row": r+2, "column": col,
                                "char": ch, "codepoint": f"U+{cp:04X}",
                                "value_excerpt": s[:160]
                            })
    return per_char, samples

def main():
    per_char, samples = scan(IN_XLSX)

    # 빈도표 저장
    freq_rows = []
    for cp, cnt in per_char.most_common():
        ch = chr(cp)
        name_hint = "REPLACEMENT CHARACTER" if cp == 0xFFFD else ""
        freq_rows.append({"codepoint": f"U+{cp:04X}", "char": ch, "count": cnt, "name_hint": name_hint})
    pd.DataFrame(freq_rows).to_csv(FREQ_CSV, index=False, encoding="utf-8-sig")

    # 샘플 저장
    sample_rows = []
    for cp, items in samples.items():
        for it in items:
            sample_rows.append(it)
    pd.DataFrame(sample_rows).to_csv(SAMPLES_CSV, index=False, encoding="utf-8-sig")

    print("[OK] filtered freq  ->", FREQ_CSV)
    print("[OK] filtered samples ->", SAMPLES_CSV)
    if freq_rows[:10]:
        print("Top suspicious:", freq_rows[:10])

if __name__ == "__main__":
    main()
