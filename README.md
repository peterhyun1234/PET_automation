# PET_automation
Title: Bonescan 대기시간 단축을 위한 자동화 서비스 개발
## 1. Research background
    - Bonescan은 뼈전이, 뼈종양, 관절염, 외상 등을 평가할 때 유용한 검사법
    - 방사성의약품 주사 후 뼈에 약품이 충분히 흡수되기까지 긴 시간이 걸림
    - 따라서 Bonescan 촬영까지 대기시간이 약 4시간 정도로 환자가 겪는 불편함이 큼 

## 2. Research goal
    - 2시간째 저화질의 뼈스캔을 추가로 촬영하여 4시간째 영상과 비교
    - 딥러닝을 통하여 저화질 영상을 고화질로 복원시키는 영상 질 개선 기술 개발 
    - 궁극적으로 뼈스캔 촬영 시 환자의 대기시간을 단축


## 3. Research procedure and method
### 3.1. Bonescan 데이터 분석 속도 개선 및 반자동화

![1](https://user-images.githubusercontent.com/46476398/92076748-003da880-edf6-11ea-9a6f-7164499889c7.JPG)

    - 뼈 스캔 분석 속도  2.7배 증진
    - Thigh and Femur ratio 분석 반자동화 (상처는 직접 판별해야 하기때문)
    - Confound를 줄이기 위한 데이터 균일화 between Thigh and Femur

### 3.2. Bonescan 데이터 노이즈 제거

![3](https://user-images.githubusercontent.com/46476398/92076743-fe73e500-edf5-11ea-8f31-77fb40a623d5.JPG)

    - 정규성 검정에 의거한 노이즈 판별 및 재측정
    - VBA를 활용해서 기존 데이터에 비해서 특이한 데이터를 추출 및 재검정

### 3.3. 판독문 수집

![2](https://user-images.githubusercontent.com/46476398/92076747-ffa51200-edf5-11ea-87f7-7023e5bf47aa.JPG)

    - 판독문에 대한 정보는 파일 형태로 따로 추출하기 어려움
    - 정규식을 통해서 필요한 정보들 파싱해서 데이터베이스에 저장

----------------------------------------    
## 4. How to process
    1. install autohotkey
    2. preparing excel data
    
### 4.1. /VBA
    - 사용자 데이터 자동 수집 및 스키마 형태로 저장
    1. Open exist excel file
    2. presss Art + F11
    3. select module for each service
    4. press F5 or F6(for debugging)
    
### 4.2. /Extract_Datas_with_bone_scan
    - 본 스캔을 통해서 데이터 자동 추출 및 노이즈 제거
    You can using this app with double click .ahk file
    Or compiring with vscode and press f5 button
    
### 4.3. /PT_review
    - 판독문 자동 크롤링 + 파싱
    You can using this app with double click .ahk file
    Or compiring with vscode and press f5 button

> 자세한건 "/04.17_filling_data/" 참조
----------------------------------------
## 5. I felt that
    * 이 연구를 통해서 많은 사람들의 불편함이 해소될 것이라는 큰 동기가 있었기 때문에 열심히 참여할 수 있었다.
    * 모든 부분에서 최고의 소프트웨어가 필요하지만 사람의 생명이랑 연관되는 의료분야에서는 더욱 최신의 최고의 소프트웨어가 필요하다는 생각이 들었다.

----------------------------------------
## 6. contact
    e-mail: peterhyun1234@gmail.com
