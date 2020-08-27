// @breif xlsx 모듈추출

const xlsx = require( "xlsx" );


// @files 엑셀 파일을 가져온다.

const excelFile = xlsx.readFile( "../data/201907_202007_data.xlsm" );

const numbersOfSample = 50;
let sheetIdx = 0;
let degreeOfMachines = 4;

Array.prototype.shuffle = function () {
    var length = this.length;
    
    // 아래에서 length 후위 감소 연산자를 사용하면서 결국 0이된다.
    // 프로그래밍에서 0은 false를 의미하기에 0이되면 종료.
    while (length) {
 
        // 랜덤한 배열 index 추출
        var index = Math.floor((length--) * Math.random());
 
        // 배열의 끝에서부터 0번째 아이템을 순차적으로 대입
        var temp = this[length];
 
        // 랜덤한 위치의 값을 맨뒤(this[length])부터 셋팅
        this[length] = this[index];
 
        // 랜덤한 위치에 위에 설정한 temp값 셋팅
        this[index] = temp;
    }
 
    // 배열을 리턴해준다.
    return this;
};


while(sheetIdx < degreeOfMachines){

    let resultGroups = [[],[],[],[],[]];
    
    // @breif 엑셀 파일의 첫번째 시트의 정보를 추출
    const sheetName = excelFile.SheetNames[sheetIdx]; 
    const currSheet = excelFile.Sheets[sheetName]; 
    
    // @details 엑셀 파일의 첫번째 시트를 읽어온다.
    console.log( "Reading sheet 1: " +  sheetName);
    const jsonData = xlsx.utils.sheet_to_json( currSheet, { defval : ""} );
    
    let idx = 1;
    let femaleAgeGroups = [[],[],[],[],[]];
    let maleAgeGroups = [[],[],[],[],[]];

    while(jsonData[idx].ID){
        // 연령별로 F/M 구분해서 배열에 저장
        //console.log(jsonData[idx]);
        if(30 <= jsonData[idx].Age && jsonData[idx].Age < 40){
            if(jsonData[idx].Sex === 'F'){
                femaleAgeGroups[0].push(jsonData[idx].ID);
            }else{
                maleAgeGroups[0].push(jsonData[idx].ID);
            }
        }else if(40 <= jsonData[idx].Age && jsonData[idx].Age < 50){
            if(jsonData[idx].Sex === 'F'){
                femaleAgeGroups[1].push(jsonData[idx].ID);
            }else{
                maleAgeGroups[1].push(jsonData[idx].ID);
            }
        }else if(50 <= jsonData[idx].Age && jsonData[idx].Age < 60){
            if(jsonData[idx].Sex === 'F'){
                femaleAgeGroups[2].push(jsonData[idx].ID);
            }else{
                maleAgeGroups[2].push(jsonData[idx].ID);
            }
        }else if(60 <= jsonData[idx].Age && jsonData[idx].Age < 70){
            if(jsonData[idx].Sex === 'F'){
                femaleAgeGroups[3].push(jsonData[idx].ID);
            }else{
                maleAgeGroups[3].push(jsonData[idx].ID);
            }
        }else if(70 <= jsonData[idx].Age && jsonData[idx].Age < 80){
            if(jsonData[idx].Sex === 'F'){
                femaleAgeGroups[4].push(jsonData[idx].ID);
            }else{
                maleAgeGroups[4].push(jsonData[idx].ID);
            }
        }
        idx++;
    }
    
    // 랜덤으로 남자 최대 50, 여자 50 넣기
    for(let i = 0; i < maleAgeGroups.length; i++){
        if(maleAgeGroups[i].length + femaleAgeGroups[i].length <= numbersOfSample*2){ // 전체 표본이 100개 이하인 경우
            while(maleAgeGroups[i].length){
                resultGroups[i].push(maleAgeGroups[i].pop());
            }
            while(femaleAgeGroups[i].length){
                resultGroups[i].push(femaleAgeGroups[i].pop());
            }
        }else{
            if(maleAgeGroups[i].length <= numbersOfSample){ // 남자가 50명 이하인 경우
                while(maleAgeGroups[i].length){
                    resultGroups[i].push(maleAgeGroups[i].pop());
                }
                // 여자는 랜덤화해서 50명 추출 후 push
                femaleAgeGroups[i].shuffle();
                let femaleCnt = 0;
                while(femaleCnt < numbersOfSample){
                    resultGroups[i].push(femaleAgeGroups[i].pop());
                    femaleCnt++;
                }

            }else{ // 남자, 여자 둘다 랜덤화해서 추출 후 push
                maleAgeGroups[i].shuffle();
                let maleCnt = 0;
                while(maleCnt < numbersOfSample){
                    resultGroups[i].push(maleAgeGroups[i].pop());
                    maleCnt++;
                }

                femaleAgeGroups[i].shuffle();
                let femaleCnt = 0;
                while(femaleCnt < numbersOfSample){
                    resultGroups[i].push(femaleAgeGroups[i].pop());
                    femaleCnt++;
                }
            }
        }

        console.log(resultGroups[i].length);
    }



    sheetIdx++; // next sheet
}


