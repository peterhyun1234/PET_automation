// @breif xlsx 모듈추출

const xlsx = require( "xlsx" );


// @files 엑셀 파일을 가져온다.
const studyTerm = "201912_202007";
const excelFile = xlsx.readFile( "../data/" + studyTerm + "_data.xlsx" );

const dataOutput = xlsx.utils.book_new();

const numbersOfSample = 25;
let sheetIdx = 0;
let degreeOfMachines = 2;

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

    let resultGroups = [];
    let resultIDs = [[],[],[],[],[]];
    
    // @breif 엑셀 파일의 첫번째 시트의 정보를 추출
    const sheetName = excelFile.SheetNames[sheetIdx]; 
    const currSheet = excelFile.Sheets[sheetName]; 
    
    console.log( "Reading sheet "+ sheetIdx +": " +  sheetName);
    
    // @details 엑셀 파일의 첫번째 시트를 읽어온다.
    let idx = 1;
    const jsonData = xlsx.utils.sheet_to_json( currSheet, { defval : ""} );
    
    let femaleAgeGroups = [[],[],[],[],[]];
    let maleAgeGroups = [[],[],[],[],[]];

    while(idx < jsonData.length){
        // 연령별로 F/M 구분해서 배열에 저장
        // console.log(jsonData[idx].ID);
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
        let age = [30, 40, 50, 60, 70];
        let femaleCnt = 0;
        let maleCnt = 0;

        if(maleAgeGroups[i].length + femaleAgeGroups[i].length <= numbersOfSample * 2){ // 전체 표본이 100개 이하인 경우
            while(maleAgeGroups[i].length){
                resultIDs[i].push(maleAgeGroups[i].pop());
                femaleCnt++;
            }
            while(femaleAgeGroups[i].length){
                resultIDs[i].push(femaleAgeGroups[i].pop());
                maleCnt++;
            }
        }else{
            
            if(maleAgeGroups[i].length <= numbersOfSample){ // 남자가 50명 이하인 경우
                while(maleAgeGroups[i].length){
                    resultIDs[i].push(maleAgeGroups[i].pop());
                    maleCnt++;
                }
                // 여자는 랜덤화해서 50명 추출 후 push
                femaleAgeGroups[i].shuffle();
                while(femaleCnt < numbersOfSample){
                    resultIDs[i].push(femaleAgeGroups[i].pop());
                    femaleCnt++;
                }

            }else{ // 남자, 여자 둘다 랜덤화해서 추출 후 push
                maleAgeGroups[i].shuffle();
                while(maleCnt < numbersOfSample){
                    resultIDs[i].push(maleAgeGroups[i].pop());
                    maleCnt++;
                }

                femaleAgeGroups[i].shuffle();
                femaleCnt = 0;
                while(femaleCnt < numbersOfSample){
                    resultIDs[i].push(femaleAgeGroups[i].pop());
                    femaleCnt++;
                }
            }
        }
        console.log("group " + age[i] + ": [ Female: " + femaleCnt + ", Male: " + maleCnt + " ]");

        for(let j = 0; j < resultIDs[i].length; j++){
            let foundInfo = jsonData.find(element => element.ID == resultIDs[i][j]);
            resultGroups.push(foundInfo);
        }

    }

    // resultGroups를 sheet별로 그대로 넣어주면 됨!


    // const currInfos = xlsx.utils.json_to_sheet(resultGroups, {skipHeader: false });

    // // @breif CELL 넓이 지정

    // currInfos["!cols"] = [
    //     { wpx: 80 }   // A열
    //     , { wpx: 80 }   // B열
    //     , { wpx: 80 }    // C열
    //     , { wpx: 60 }    // D열
    //     , { wpx: 60 }    // D열
    // ]



    //xlsx.utils.book_append_sheet(dataOutput, currInfos, sheetName);

    
    sheetIdx++; // next sheet
}

// files 엑셀파일을 생성하고 저장한다.
//xlsx.writeFile(dataOutput, studyTerm + "_mined_data.xlsx"); 

