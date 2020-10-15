# confounds 줄이려면 
1. 성별도 비슷하게 맞춰야하지 않나? M:F = 1:1로
2. 측정 날짜도 랜덤으로 뽑아낼 필요가 있음

# 구현 방향
1. 각 기계에서 M, F 별 나잇대별 50개씩 뽑기
2. 일단 30대-70대 데이터 100개씩 추출

# 1. 데이터 정제 및 추출
	"201912_202007_data.xlsm" 에서 2000개 목표

	by using "readAndPreprocessing.js" with node.js

	1. Infi 474개
	2. NMCT670 473개
	3. Sym 476개
	4. NM830 375개 (30대, 70대 부족 함)

# 2. Classfication
	by using "check_sharpness.ahk"

# 3. Insert sharpness data
	into "201912_202007_mined_data.xlsx"
	from "201912_202007_result.txt"

# 4. Preprocessing in high order of sharpness

	by using "PreprocessingHighSharpness.js"

	to "201912_202007_preprocessed_data.xlsx"
	from "201912_202007_mined_data.xlsx" 
<br>

	[ 추출할 표본 갯수 = 50 ]
	[ 기계 = Infi ]
	[ 연령층 = 30대 ]

	0. Select (기계) sheet
	1. (기계) sheet에서 (연령층) 데이터를 성별 별로 배열에 추가한다.
	2. 성별 별 데이터의 갯수가 (추출할 표본 갯수)이하일 경우 
		2.1. (연령층) 데이터의 성별별 데이터를 추가한다. 
	3. 성별 별 데이터의 갯수가 (추출할 표본 갯수)이상 일 경우
		3.1. 랜덤으로 50개만 추가한다. 
	4. 연령층 += 10대
	5. 70대까지 반복한다.
	6. 기계 = next_기계

# 5. Data parsing and analysis

	by using "orderPreprocessedData.cls"

	to "2019012_202007_preprocessed_FTTR.xlsm"
	from "201912_202007_preprocessed_data.xlsx" 
<br>

