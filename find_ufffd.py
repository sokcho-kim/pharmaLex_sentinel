import pandas as pd
import os
from collections import Counter

# 파일 경로
excel_path = r"C:\Jimin\pharmaLex_sentinel\data\요양심사약제_후처리.xlsx"

# Excel 전체 시트 불러오기
xls = pd.ExcelFile(excel_path)

report = []  # 결과 담을 리스트
char_counter = Counter()

for sheet in xls.sheet_names:
    df = pd.read_excel(excel_path, sheet_name=sheet, dtype=str)  # 문자열로 읽기
    for col in df.columns:
        for idx, val in df[col].items():
            if pd.isna(val):
                continue
            if "�" in val:  # replacement character 포함 여부
                count = val.count("�")
                char_counter["�"] += count
                report.append({
                    "sheet": sheet,
                    "row": idx + 2,   # 엑셀은 보통 헤더 포함하므로 +2
                    "column": col,
                    "value": val,
                    "count_in_cell": count
                })

# 전체 빈도 요약
print("총 발견 건수:", len(report))
print("문자별 빈도:", dict(char_counter))

# 상위 20건 미리보기
pd.set_option("display.max_colwidth", 100)
print(pd.DataFrame(report).head(20))

# CSV로 저장 (전체 목록)
out_csv = os.path.join(os.path.dirname(excel_path), "invalid_char_report.csv")
pd.DataFrame(report).to_csv(out_csv, index=False, encoding="utf-8-sig")
print("리포트 저장 완료:", out_csv)
