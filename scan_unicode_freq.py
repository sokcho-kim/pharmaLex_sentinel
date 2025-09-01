# -*- coding: utf-8 -*-
import os, pandas as pd
from collections import Counter, defaultdict

BASE = r"C:\Jimin\pharmaLex_sentinel"
IN_XLSX = os.path.join(BASE, r"data\요양심사약제_후처리.xlsx")
OUT_DIR = os.path.join(BASE, "out_scan")
os.makedirs(OUT_DIR, exist_ok=True)

FREQ_CSV    = os.path.join(OUT_DIR, "unicode_freq.csv")
SAMPLES_CSV = os.path.join(OUT_DIR, "unicode_samples.csv")
CELLS_CSV   = os.path.join(OUT_DIR, "unicode_in_cells.csv")

# 스캔 대상: 비ASCII 전체(>0x7F) + ASCII라도 의심문자 목록(옵션)
SUSPECT_ASCII = set("?")  # 필요 없으면 빈 set()

def scan_workbook(path):
    xls = pd.ExcelFile(path)
    per_char = Counter()
    per_cell_rows = []
    samples = defaultdict(list)

    for sheet in xls.sheet_names:  # 시트 한 개여도 일반화
        df = pd.read_excel(path, sheet_name=sheet, dtype=str)
        for r_idx in range(len(df)):
            for col in df.columns:
                val = df.iat[r_idx, df.columns.get_loc(col)]
                if pd.isna(val):
                    continue
                s = str(val)

                # 셀 단위 문자 카운트
                cell_counter = Counter()
                for ch in s:
                    cp = ord(ch)
                    if cp > 127 or ch in SUSPECT_ASCII:
                        per_char[cp] += 1
                        cell_counter[cp] += 1
                        # 샘플은 문자별 최대 5건만 저장
                        if len(samples[cp]) < 5:
                            samples[cp].append({
                                "sheet": sheet,
                                "row": r_idx + 2,  # 엑셀표시행
                                "column": col,
                                "value_excerpt": s[:160]
                            })

                if cell_counter:
                    per_cell_rows.append({
                        "sheet": sheet,
                        "row": r_idx + 2,
                        "column": col,
                        "value": s,
                        "char_count": sum(cell_counter.values()),
                        "codepoints": ";".join(f"U+{cp:04X}×{cnt}" for cp, cnt in sorted(cell_counter.items()))
                    })

    return per_char, samples, per_cell_rows

def main():
    per_char, samples, per_cell_rows = scan_workbook(IN_XLSX)

    # 1) 문자 빈도표
    freq_rows = []
    for cp, cnt in per_char.most_common():
        ch = chr(cp)
        freq_rows.append({
            "codepoint": f"U+{cp:04X}",
            "char": ch,
            "count": cnt,
            "name_hint": (
                "REPLACEMENT CHARACTER" if cp == 0xFFFD else ""
            )
        })
    pd.DataFrame(freq_rows).to_csv(FREQ_CSV, index=False, encoding="utf-8-sig")

    # 2) 문자별 샘플
    sample_rows = []
    for cp, items in samples.items():
        for it in items:
            it2 = {
                "codepoint": f"U+{cp:04X}",
                "char": chr(cp),
                **it
            }
            sample_rows.append(it2)
    pd.DataFrame(sample_rows).to_csv(SAMPLES_CSV, index=False, encoding="utf-8-sig")

    # 3) 셀 상세
    pd.DataFrame(per_cell_rows).to_csv(CELLS_CSV, index=False, encoding="utf-8-sig")

    print("[OK] 문자 빈도표 :", FREQ_CSV)
    print("[OK] 문자 샘플   :", SAMPLES_CSV)
    print("[OK] 셀 상세표   :", CELLS_CSV)
    # 콘솔 요약 상위 15개
    print("\nTop 15 codepoints:")
    for row in freq_rows[:15]:
        print(f"{row['codepoint']} ({row['char']}) -> {row['count']}")

if __name__ == "__main__":
    main()
