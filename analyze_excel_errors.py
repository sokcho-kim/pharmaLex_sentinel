import pandas as pd
import re
import json
from pathlib import Path

def load_excel_data():
    """요양심사약제 후처리 엑셀 파일 로드"""
    file_path = 'data/요양심사약제_후처리.xlsx'
    try:
        # 모든 시트 읽기
        excel_file = pd.ExcelFile(file_path)
        print(f"Excel 파일의 시트: {excel_file.sheet_names}")
        
        # 첫 번째 시트 데이터 로드
        df = pd.read_excel(file_path, sheet_name=0)
        print(f"첫 번째 시트 크기: {df.shape}")
        print(f"컬럼명: {list(df.columns)}")
        
        return df, excel_file.sheet_names
    except Exception as e:
        print(f"Excel 파일 로드 오류: {e}")
        return None, []

def load_ocr_patterns():
    """OCR 오류 패턴 로드"""
    ocr_df = pd.read_csv('out/ocr_unit_anomalies_scan.csv')
    
    # 오류 패턴 매핑 생성
    error_patterns = {}
    
    for idx, row in ocr_df.iterrows():
        match = row['match']
        suggested_fix = row['suggested_fix']
        classification = row['classification']
        
        # 자동 교정 가능한 패턴들
        if 'mcg' in match:
            error_patterns['mcg'] = '㎍'
        elif 'ug' in match and 'mcg' not in match:
            error_patterns['ug'] = '㎍'
        elif classification == 'suspect_micro_as_g':
            # 소량 g → ㎍ 패턴
            numbers = re.findall(r'(\d+\.?\d*)', match)
            if numbers:
                value = float(numbers[0])
                if 1 <= value <= 100:
                    # 특정 패턴들을 등록
                    error_patterns[match] = suggested_fix
    
    return error_patterns

def detect_unit_errors_in_excel(df, error_patterns):
    """엑셀 데이터에서 단위 오류 탐지"""
    errors_found = []
    
    # 모든 텍스트 컬럼에 대해 검사
    text_columns = df.select_dtypes(include=['object']).columns
    
    for col in text_columns:
        for idx, cell_value in enumerate(df[col]):
            if pd.isna(cell_value):
                continue
                
            cell_str = str(cell_value)
            
            # 각 오류 패턴에 대해 검사
            for error_pattern, correct_pattern in error_patterns.items():
                if error_pattern in cell_str:
                    # 오류 발견
                    corrected_value = cell_str.replace(error_pattern, correct_pattern)
                    
                    errors_found.append({
                        'row_idx': idx,
                        'column': col,
                        'original_value': cell_str,
                        'error_pattern': error_pattern,
                        'corrected_value': corrected_value,
                        'correction': correct_pattern
                    })
            
            # 추가적인 일반적인 단위 오류 패턴들 검사
            # g → ㎍ 오류 (제형 단어와 함께)
            form_keywords = ['정', '주', '시럽', '이식제', '캡슐', '패치', '외용제', '점안액', '연고', '주사']
            
            # 숫자+g 패턴과 제형 단어 검사
            g_matches = re.findall(r'(\d+\.?\d*)\s*g(?![a-zA-Z])', cell_str)
            if g_matches and any(keyword in cell_str for keyword in form_keywords):
                for g_match in g_matches:
                    value = float(g_match)
                    if 1 <= value <= 100:  # 소량이면서 제형 단어가 있으면
                        original_pattern = f"{g_match}g"
                        corrected_pattern = f"{g_match}㎍"
                        corrected_value = cell_str.replace(original_pattern, corrected_pattern)
                        
                        errors_found.append({
                            'row_idx': idx,
                            'column': col,
                            'original_value': cell_str,
                            'error_pattern': original_pattern,
                            'corrected_value': corrected_value,
                            'correction': corrected_pattern
                        })
    
    return errors_found

def create_corrected_dataframe(df, errors_found):
    """오류를 수정한 새로운 데이터프레임 생성"""
    df_corrected = df.copy()
    
    # 각 오류에 대해 수정 적용
    for error in errors_found:
        row_idx = error['row_idx']
        col = error['column']
        corrected_value = error['corrected_value']
        
        df_corrected.at[row_idx, col] = corrected_value
    
    return df_corrected

def save_error_log(errors_found):
    """오류 수정 내역을 CSV로 저장"""
    if errors_found:
        error_df = pd.DataFrame(errors_found)
        error_df.to_csv('out/error_corrections.csv', index=False, encoding='utf-8-sig')
        print(f"오류 수정 내역이 out/error_corrections.csv에 저장되었습니다.")
        return error_df
    else:
        print("발견된 오류가 없습니다.")
        return None

def main():
    print("=== 요양심사약제 데이터 오류 분석 및 수정 ===\n")
    
    # 1. 엑셀 데이터 로드
    print("1. 엑셀 데이터 로드 중...")
    df, sheet_names = load_excel_data()
    if df is None:
        return
    
    # 2. OCR 오류 패턴 로드
    print("\n2. OCR 오류 패턴 분석 중...")
    error_patterns = load_ocr_patterns()
    print(f"로드된 오류 패턴: {len(error_patterns)}개")
    for pattern, correction in list(error_patterns.items())[:5]:
        print(f"  {pattern} → {correction}")
    
    # 3. 엑셀 데이터에서 오류 탐지
    print("\n3. 엑셀 데이터에서 단위 오류 탐지 중...")
    errors_found = detect_unit_errors_in_excel(df, error_patterns)
    
    print(f"\n=== 탐지 결과 ===")
    print(f"총 발견된 오류: {len(errors_found)}개")
    
    if errors_found:
        # 오류 요약 출력
        print("\n발견된 오류들:")
        for i, error in enumerate(errors_found[:10]):  # 처음 10개만 출력
            print(f"{i+1}. 행 {error['row_idx']}, 컬럼 '{error['column']}'")
            print(f"   원본: {error['original_value']}")
            print(f"   수정: {error['corrected_value']}")
            print(f"   패턴: {error['error_pattern']} → {error['correction']}\n")
        
        if len(errors_found) > 10:
            print(f"... 외 {len(errors_found)-10}개 더")
        
        # 4. 수정된 데이터프레임 생성
        print("\n4. 수정된 데이터 생성 중...")
        df_corrected = create_corrected_dataframe(df, errors_found)
        
        # 5. 수정된 파일 저장
        corrected_file_path = 'out/요양심사약제_후처리_수정본.xlsx'
        df_corrected.to_excel(corrected_file_path, index=False)
        print(f"수정된 파일이 {corrected_file_path}에 저장되었습니다.")
        
        # 6. 오류 수정 로그 저장
        print("\n5. 오류 수정 내역 저장 중...")
        error_log = save_error_log(errors_found)
        
    else:
        print("단위 오류가 발견되지 않았습니다.")

if __name__ == "__main__":
    main()