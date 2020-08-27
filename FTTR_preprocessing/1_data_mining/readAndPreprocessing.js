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
    while(jsonData[idx].ID){
        //console.log( jsonData[idx].ID);
        idx++;
    }

    console.log(idx);

    sheetIdx++; // next sheet
}
