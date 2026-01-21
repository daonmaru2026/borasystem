import pandas as pd
import os
import re
import glob

# [1단계] 파일 자동 스캔 설정
# 바탕화면 경로 자동 인식
user_profile = os.environ['USERPROFILE']
base_path = os.path.join(user_profile, 'OneDrive', '바탕 화면')
if not os.path.exists(base_path): # 원드라이브 없으면 그냥 바탕화면
    base_path = os.path.join(user_profile, 'Desktop')

# 파일명에 'y'가 들어가는 엑셀 파일은 모두 찾기 (예: 2022y.xlsx, 2026y.xlsx 등)
target_pattern = os.path.join(base_path, '*y.xlsx')
files = glob.glob(target_pattern)

print(f"📂 검색 경로: {base_path}")
print(f"🔎 발견된 연도별 파일: {len(files)}개")

def final_refine_logic(text):
    if pd.isna(text) or str(text).strip() == "": return "삭제대상"
    t = str(text).replace(' ', '').upper()
    
    # 0. 소형 우선 분류
    if '/다' in t or '다마' in t: return "다마스"
    if '/라' in t or '라보' in t: return "라보"
    if '/오' in t or '오토' in t: return "오토바이"

    # 1. 톤수 추출
    ton = ""
    if '2.5' in t or '25톤' in t or t.startswith('2.5'): ton = "2.5톤"
    elif '3.5' in t or '35' in t: ton = "3.5톤"
    elif '5톤' in t or '5T' in t or '5축' in t or '5톤축' in t: ton = "5톤"
    elif any(k in t for k in ['1.4', '1.3', '1.5']): ton = "1.4톤"
    elif any(k in t for k in ['1톤', '1T', '1카', '1탑', '1윙']): ton = "1톤"
    elif any(x in t for x in ['11', '16', '25']) and '톤' in t:
        m = re.search(r'(\d+)톤', t)
        ton = m.group(0) if m else "대형"
    else:
        p_match = re.search(r'(\d+)P', t)
        if p_match: return f"{p_match.group(1)}P"
        return "미분류"

    # 2. 옵션 판별
    is_lift = any(k in t for k in ['리프트', '리프', '리', 'LIFT'])
    is_wing_top = any(k in t for k in ['윙', '탑', 'WING', 'TOP', '캅'])
    is_wide = '광폭' in t or '광' in t
    is_axis = '축' in t
    is_no_vibe = '무진동' in t

    # 3. 명칭 확정
    if ton == "5톤" and is_axis:
        res = "5톤축차"
    else:
        res = ton
    
    if is_wide and ton not in ["1톤", "1.4톤", "2.5톤"] and res != "5톤축차":
        res += "광폭"
    
    if is_no_vibe: res += "/무진동"
    if is_lift: res += "리프트"
    elif is_wing_top: res += "탑/윙"
    
    return res

try:
    all_data = []
    if not files:
        print("❌ '20xx.xlsx' 형식의 파일을 찾을 수 없습니다.")
    else:
        for full_p in files:
            f_name = os.path.basename(full_p)
            print(f"📦 {f_name} 통합 중...")
            try:
                tmp = pd.read_excel(full_p)
                all_data.append(tmp)
            except Exception as e:
                print(f"⚠️ {f_name} 읽기 실패: {e}")

    if all_data:
        df = pd.concat(all_data, ignore_index=True)
        df['배달운임'] = df['배달운임'].astype(str).str.replace(',', '').str.extract(r'(\d+)').astype(float).fillna(0)
        df['접수일자'] = pd.to_datetime(df['접수일자'], errors='coerce').dt.strftime('%y/%m/%d')
        df['차종_최종'] = df['도 착 지'].apply(final_refine_logic)
        df = df[df['차종_최종'] != "삭제대상"]
        df = df.sort_values(by='접수일자', ascending=False)
        output_p = os.path.join(base_path, '보라물류_최종정밀단가표.xlsx')
        
        save_cols = ['접수일자', '고객성명', '도 착 지', '차종_최종', '배달운임']
        real_cols = [c for c in save_cols if c in df.columns]
        
        df[real_cols].to_excel(output_p, index=False)
        print("\n🚀 [성공] '보라물류_최종정밀단가표.xlsx' 생성 완료!")
        print(f"저장 위치: {output_p}")
        
except Exception as e:
    print(f"❌ 오류 발생: {e}")
    input("엔터를 누르면 종료합니다...")
