# PULSE TUNE 도우미

import pandas as pd

# 엑셀 파일 경로
file_path = r'C:\Users\whs13\OneDrive\Desktop\파이썬\240317_test(2).xlsx'

# 엑셀 파일 읽어오기
df = pd.read_excel(file_path, header=None)

# 문자열 값을 숫자로 변환하여 숫자 데이터만 남기기
numeric_df = df.apply(pd.to_numeric, errors='coerce')

# 특정 행부터 마지막 행까지의 평균값, 최댓값, 최솟값을 계산
start_row = 1  # 첫 번째 행부터 시작
end_row = numeric_df.shape[0] - 1  # 마지막 행 인덱스 가져오기
mean_values = numeric_df.iloc[start_row:].mean()
max_values = numeric_df.iloc[start_row:].max()
min_values = numeric_df.iloc[start_row:].min()

# 새로운 행으로 추가
df.loc[end_row + 1] = mean_values
df.loc[end_row + 2] = max_values
df.loc[end_row + 3] = min_values

# 사용자 입력 값을 받기
threshold = float(input("몇 이상을 1로 잡을지 정해주세요: "))

# 평균값이 사용자 입력 값보다 크면 1, 그렇지 않으면 0을 할당
df.loc[end_row + 4] = [1 if value > threshold else 0 for value in mean_values]

# ExcelWriter 객체 생성
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    # 수정된 데이터프레임을 새로운 시트로 저장
    df.to_excel(writer, sheet_name='새로운 시트', index=False, header=False)

print(f"'{file_path}' 파일에 새로운 시트로 데이터가 저장되었습니다.")