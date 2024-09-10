import pandas as pd
from datetime import datetime

# 두 개의 엑셀 파일을 불러옵니다. (혹은 동일 파일 내 시트를 불러올 수도 있습니다.)
file_path = "/Users/bbw500/Downloads/sf_list.xlsx"  # 엑셀 파일 경로
sheet1 = pd.read_excel(file_path, sheet_name='sf')  # 첫 번째 시트
sheet2 = pd.read_excel(file_path, sheet_name='cms')  # 두 번째 시트
print(sheet2)

# Sheet1의 특정 셀 값(A 열)과 Sheet2에서 해당 값을 찾아 비교
# def find_values_in_sheet2(sheet1, sheet2):
#     # Sheet1의 D컬럼 값을 기준으로 Sheet2의 A컬럼에 해당 값이 있는지 여부를 확인하고, 그 결과를 새로운 열에 추가
#     sheet2['세일즈포스존재'] = sheet2['계약번호'].apply(lambda x: '있음' if x in sheet1['계약 번호'].values else '없음')
#
# # 함수 실행
# find_values_in_sheet2(sheet1, sheet2)
#
# # 결과 확인 (필요 시 파일로 저장)
# print(sheet2[['계약번호', '세일즈포스존재']])


def compare_sheets(sheet1, sheet2):
    # Sheet2의 A열 값을 기준으로 Sheet1에서 찾고, Sheet1의 같은 행의 B열 값을 가져옵니다.
    sheet2['Comparison'] = sheet2.apply(lambda row: compare_value(row['계약번호'], row['계약해지일'], sheet2), axis=1)


def compare_value(value_in_sheet2_A, value_in_sheet2_B, sheet2):
    # Sheet1에서 A열에 해당하는 값이 있는지 찾습니다.
    matching_row = sheet1[sheet1['계약 번호'] == value_in_sheet2_A]

    if not matching_row.empty:
        # print(f"계약번호 : {value_in_sheet2_A} - {matching_row['계약 번호'].values[0]} - {matching_row['해지일'].values[0]}")
        # 일치하는 값이 있으면 해당 행의 B열 값을 비교합니다.
        value_in_sheet1_B = matching_row['해지일'].values[0]
        date_obj = datetime.strptime(value_in_sheet1_B, "%Y. %m. %d.")
        if value_in_sheet2_B == date_obj:
            return "일치"
        else:
            return "불일치"
    else:
        # Sheet2에 해당하는 값이 없을 경우
        return "Sheet2에 없음"


# 비교 함수 실행
compare_sheets(sheet1, sheet2)



# 비교 함수 실행
# compare_sheets(sheet1, sheet2)
#
# # 결과 확인 (필요 시 파일로 저장)
print(sheet2[['계약번호', '계약해지일', 'Comparison']])
# 결과를 새로운 엑셀 파일로 저장
output_file = "/Users/bbw500/Downloads/find_values_result.xlsx"
sheet2.to_excel(output_file, index=False)
# 결과를 새로운 엑셀 파일로 저장