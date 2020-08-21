# PET_automation
Title: 뼈스캔 대기시간 단축을 위한 딥러닝 기반의 고화질 영상 복원 인공지능 알고리즘 개발
------------------------------------
## 1. Research background
    - 뼈스캔은 뼈전이, 뼈종양, 관절염, 외상 등을 평가할 때 유용한 검사법이나, 방사성의약품 주사 후 뼈에 약품이 충분히 흡수되기까지 긴 시간이 걸림
    - 따라서 뼈스캔 촬영까지 환자의 대기시간이 약 4시간 정도로 길어서 환자가 겪는 불편함이 큼 

## 2. Research goal
    - 2시간째 저화질의 뼈스캔을 추가로 촬영하여 4시간째 영상과 비교하고, 인공지능을 이용한 딥러닝을 통하여 저화질 영상을 고화질로 복원시키는 영상 질 개선 기술을 개발 
    - 이러한 영상 질 개선을 통하여 궁극적으로 뼈스캔 촬영 시 환자의 대기시간을 단축
---------------------------------------- 


## 3. Research procedure and method
### 3.1. 뼈 스캔 데이터 추출 속도 개선 및 자동화
    (사진 1)
    - 다른 기계의 스토리지를 원격으로 접근해서 사용
    - 로컬 스토리지의 공간의 한계를 극복
    - 원격 접근이기 때문에 네트워크 전송이 중요함
    - 데이터 복제 저장을 통한 결함 감래 시스템
    - 복제할 때 병목현상을 방지하기 위한 릴레이 전송  

### 3.2. 뼈 스캔 데이터 노이즈 제거
    (사진 2)
    - 다른 기계의 스토리지를 원격으로 접근해서 사용
    - 로컬 스토리지의 공간의 한계를 극복
    - 원격 접근이기 때문에 네트워크 전송이 중요함
    - 데이터 복제 저장을 통한 결함 감래 시스템
    - 복제할 때 병목현상을 방지하기 위한 릴레이 전송  

### 3.2. 분산 스토리지를 통한 데이터 저장
    (사진 3)
    - 다른 기계의 스토리지를 원격으로 접근해서 사용
    - 로컬 스토리지의 공간의 한계를 극복
    - 데이터 복제 저장을 통한 fault-tolerance 시스템
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
