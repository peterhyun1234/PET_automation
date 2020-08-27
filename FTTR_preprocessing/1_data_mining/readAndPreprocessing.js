// @breif xlsx 모듈추출

const xlsx = require( "xlsx" );


// @files 엑셀 파일을 가져온다.

const excelFile = xlsx.readFile( "../data/201907_202007_data.xlsm" );

const numbersOfSample = 50;
let sheetIdx = 0;
let degreeOfMachines = 4;

while(sheetIdx < degreeOfMachines){
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
    
    for(let i = 0; i < maleAgeGroups.length; i++){
        console.log(maleAgeGroups[i].length);
    }


    sheetIdx++; // next sheet
}
