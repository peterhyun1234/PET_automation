// @breif xlsx 모듈추출

const xlsx = require( "xlsx" );


// @files 엑셀 파일을 가져온다.
const studyTerm = "201912_202007";
const excelFile = xlsx.readFile( "../data/1_result/" + studyTerm + "/" + studyTerm + "_mined_data.xlsx" );

const dataOutput = xlsx.utils.book_new();

const numbersOfSample = 50;
let sheetIdx = 0;
let degreeOfMachines = 4;


while(sheetIdx < degreeOfMachines){

    let resultGroups = [];
    let resultIDs = [[],[],[],[],[]];
    
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
    console.log(jsonData.length);
    

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

