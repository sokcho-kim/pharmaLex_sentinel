import pandas as pd
import re
import json
import sys
import io

# Windows 환경에서 UTF-8 출력 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def load_data():
    """OCR 단위 이상 탐지 CSV 파일 로드"""
    df = pd.read_csv('out/ocr_unit_anomalies_scan.csv')
    return df

def classify_unit_corrections(df):
    """단위 교정을 자동 교정 가능 / 사람 검토 필요로 분류"""
    auto_corrections = []
    manual_reviews = []
    
    for idx, row in df.iterrows():
        match = row['match']
        classification = row['classification']
        context = row['context']
        
        # 자동 교정 가능한 케이스들
        if 'ug' in match or 'mcg' in match:
            # ASCII 단위를 유니코드 기호로 변환
            if 'ug' in match and 'mcg' not in match:
                corrected = match.replace('ug', '㎍')
                auto_corrections.append({
                    'original': match,
                    'corrected': corrected,
                    'reason': 'ASCII ug → 기호 ㎍',
                    'context': context[:50] + '...' if len(context) > 50 else context
                })
            elif 'mcg' in match:
                corrected = match.replace('mcg', '㎍')
                auto_corrections.append({
                    'original': match,
                    'corrected': corrected,
                    'reason': 'ASCII mcg → 기호 ㎍',
                    'context': context[:50] + '...' if len(context) > 50 else context
                })
        
        # 소량 g 단위 분석 (1~100g 범위)
        elif classification == 'suspect_micro_as_g':
            # 숫자 추출
            numbers = re.findall(r'(\d+\.?\d*)', match)
            if numbers:
                value = float(numbers[0])
                
                # 제형 단어 확인
                form_keywords = ['정', '주', '시럽', '이식제', '캡슐', '패치', '외용제', '점안액', '연고', '주사']
                has_form_hint = any(keyword in context for keyword in form_keywords)
                
                if 1 <= value <= 100 and has_form_hint:
                    # 자동 교정 가능
                    corrected = match.replace('g', '㎍')
                    auto_corrections.append({
                        'original': match,
                        'corrected': corrected,
                        'reason': f'소량 g({value}g), 제형 단어 있음 → ㎍',
                        'context': context[:50] + '...' if len(context) > 50 else context
                    })
                else:
                    # 사람 검토 필요
                    if value > 100:
                        reason = f'큰 값({value}g) → 실제 g일 가능성'
                    else:
                        reason = f'소량 g({value}g)이나 제형 단어 없음'
                    
                    manual_reviews.append({
                        'original': match,
                        'suggested': match.replace('g', '㎍(검토)'),
                        'reason': reason,
                        'context': context[:80] + '...' if len(context) > 80 else context
                    })
        
        # 기타 사람 검토 필요한 케이스들
        elif classification in ['review_micro_as_g', 'suspect_greek_broken', 'suspect_mu_alone']:
            if classification == 'suspect_mu_alone':
                reason = 'μ 단독 검출 → 맥락 검토 필요 (㎍/㎖?)'
                suggested = '㎍/㎖?'
            elif classification == 'suspect_greek_broken':
                reason = 'α,β,γ 등 그리스 문자 오인 가능성'
                suggested = '그리스 문자?'
            else:
                reason = '기타 검토 필요'
                suggested = '검토 필요'
            
            manual_reviews.append({
                'original': match,
                'suggested': suggested,
                'reason': reason,
                'context': context[:80] + '...' if len(context) > 80 else context
            })
    
    return auto_corrections, manual_reviews

def generate_correction_rules(auto_corrections):
    """일괄 교정 규칙 생성"""
    rules = set()
    
    for correction in auto_corrections:
        original = correction['original']
        corrected = correction['corrected']
        
        # 패턴 기반 규칙 추출
        if 'ug' in original and 'mcg' not in original:
            rules.add('"ug" → "㎍"')
        elif 'mcg' in original:
            rules.add('"mcg" → "㎍"')
        elif 'g' in original and '㎍' in corrected:
            # 숫자 범위 기반 규칙
            numbers = re.findall(r'(\d+\.?\d*)', original)
            if numbers:
                value = float(numbers[0])
                if 1 <= value <= 100:
                    rules.add('"1~100g + 제형 단어" → "㎍"')
    
    return sorted(list(rules))

def generate_mapping_json(auto_corrections):
    """최종 매핑 JSON 사전 생성"""
    mapping = {}
    
    for correction in auto_corrections:
        original = correction['original']
        corrected = correction['corrected']
        
        # 일반화된 패턴으로 매핑
        if 'ug' in original and 'mcg' not in original:
            mapping['ug'] = '㎍'
        elif 'mcg' in original:
            mapping['mcg'] = '㎍'
        elif original.endswith('g') and corrected.endswith('㎍'):
            # 구체적인 숫자+g 패턴은 개별적으로 처리하지 않고
            # 조건부 규칙으로 처리 (코드에서 구현 필요)
            mapping[' g'] = ' ㎍'  # 일반적인 g → ㎍ 패턴
    
    return mapping

def print_report(auto_corrections, manual_reviews, rules, mapping):
    """최종 리포트 출력"""
    print("# OCR 단위/기호 정규화 검증 리포트")
    print("\n## ✅ 자동 교정 목록")
    print("| 원본 match | 교정 제안 | 사유 |")
    print("|------------|-----------|------|")
    
    for correction in auto_corrections:
        print(f"| {correction['original']} | {correction['corrected']} | {correction['reason']} |")
    
    print(f"\n**자동 교정 총 건수: {len(auto_corrections)}건**")
    
    print("\n## ⚠️ 사람 검토 필요 목록")
    print("| 원본 match | 교정 제안 | 사유 | context |")
    print("|------------|-----------|------|---------|")
    
    for review in manual_reviews:
        print(f"| {review['original']} | {review['suggested']} | {review['reason']} | \"{review['context']}\" |")
    
    print(f"\n**검토 필요 총 건수: {len(manual_reviews)}건**")
    
    print("\n## 📌 일괄 교정 규칙")
    for rule in rules:
        print(f"- {rule}")
    
    print("\n## 🗂 최종 매핑 JSON")
    print("```json")
    print(json.dumps(mapping, ensure_ascii=False, indent=2))
    print("```")

def main():
    # 데이터 로드
    df = load_data()
    print(f"총 {len(df)}개 레코드 로드")
    
    # 분류 실행
    auto_corrections, manual_reviews = classify_unit_corrections(df)
    
    # 규칙 생성
    rules = generate_correction_rules(auto_corrections)
    mapping = generate_mapping_json(auto_corrections)
    
    # 리포트 출력
    print_report(auto_corrections, manual_reviews, rules, mapping)

if __name__ == "__main__":
    main()