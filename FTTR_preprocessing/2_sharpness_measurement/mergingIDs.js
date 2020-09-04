const xlsx = require( "xlsx" );
const fs = require('fs');

const studyTerm = "201912_202007";
const excelFile = xlsx.readFile( "../1_data_mining/result/" + studyTerm + "/" + studyTerm + "_mined_data.xlsx" );
let sheetIdx = 0;
let degreeOfMachines = 4;

while(sheetIdx < degreeOfMachines){
    
    const sheetName = excelFile.SheetNames[sheetIdx]; 
    const currSheet = excelFile.Sheets[sheetName]; 

    console.log( "Reading sheet 1: " +  sheetName);
    const jsonData = xlsx.utils.sheet_to_json( currSheet, { defval : ""} );
    
    let data_len = jsonData.length;
    let outputString = "";
    
    for(let i = 0; i < data_len; i++){
        if(i === data_len-1){
            outputString += jsonData[i].ID;
        }else{
            outputString += jsonData[i].ID + ", ";
        }
    }
    console.log(outputString);
    //console.log(data_len);
    //fs.writeFileSync("merged_IDs.txt", '\ufeff' + outputString, {encoding: 'utf8'});
    
    sheetIdx++; // next sheet
}
