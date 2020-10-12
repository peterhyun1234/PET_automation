// @breif xlsx 모듈추출

const xlsx = require("xlsx");


// @files 엑셀 파일을 가져온다.
const studyTerm = "201912_202007";
const excelFile = xlsx.readFile("../data/1_result/" + studyTerm + "/" + studyTerm + "_mined_data.xlsx");

const dataOutput = xlsx.utils.book_new();

const numbersOfSample = 50;
let sheetIdx = 0;
let degreeOfMachines = 4;


while (sheetIdx < degreeOfMachines) {

    
    // @breif 엑셀 파일의 첫번째 시트의 정보를 추출
    const sheetName = excelFile.SheetNames[sheetIdx];
    const currSheet = excelFile.Sheets[sheetName];
    
    // @details 엑셀 파일의 첫번째 시트를 읽어온다.
    console.log("[Reading sheet " + sheetIdx + ": " + sheetName + "]");
    const jsonData = xlsx.utils.sheet_to_json(currSheet, { defval: "" });

    let idx = 0;
    let sumOfSharpness = 0;
    let femaleAgeGroups = [[], [], [], [], []];
    let maleAgeGroups = [[], [], [], [], []];
    let resultIDs = [[], [], [], [], []];
    
    while (idx < jsonData.length) {
        //console.log(jsonData[idx].ID);

        sumOfSharpness += jsonData[idx].Sharpeness;
        // 연령별로 F/M 구분해서 배열에 저장
        if (30 <= jsonData[idx].Age && jsonData[idx].Age < 40) {
            if (jsonData[idx].Sex === 'F') {
                femaleAgeGroups[0].push([jsonData[idx].ID, jsonData[idx].Sharpeness]);
            } else {
                maleAgeGroups[0].push([jsonData[idx].ID, jsonData[idx].Sharpeness]);
            }
        } else if (40 <= jsonData[idx].Age && jsonData[idx].Age < 50) {
            if (jsonData[idx].Sex === 'F') {
                femaleAgeGroups[1].push([jsonData[idx].ID, jsonData[idx].Sharpeness]);
            } else {
                maleAgeGroups[1].push([jsonData[idx].ID, jsonData[idx].Sharpeness]);
            }
        } else if (50 <= jsonData[idx].Age && jsonData[idx].Age < 60) {
            if (jsonData[idx].Sex === 'F') {
                femaleAgeGroups[2].push([jsonData[idx].ID, jsonData[idx].Sharpeness]);
            } else {
                maleAgeGroups[2].push([jsonData[idx].ID, jsonData[idx].Sharpeness]);
            }
        } else if (60 <= jsonData[idx].Age && jsonData[idx].Age < 70) {
            if (jsonData[idx].Sex === 'F') {
                femaleAgeGroups[3].push([jsonData[idx].ID, jsonData[idx].Sharpeness]);
            } else {
                maleAgeGroups[3].push([jsonData[idx].ID, jsonData[idx].Sharpeness]);
            }
        } else if (70 <= jsonData[idx].Age && jsonData[idx].Age < 80) {
            if (jsonData[idx].Sex === 'F') {
                femaleAgeGroups[4].push([jsonData[idx].ID, jsonData[idx].Sharpeness]);
            } else {
                maleAgeGroups[4].push([jsonData[idx].ID, jsonData[idx].Sharpeness]);
            }
        }
        idx++;
    }
    
    console.log("Number of samples: " + idx);

    console.log("");

    let curAgeRange = 30;
    for (let i = 0; i < 5; i++) {
        femaleAgeGroups[i].sort(function(a, b) {
            return b[1] - a[1];
        });
        maleAgeGroups[i].sort(function(a, b) {
            return b[1] - a[1];
        });
        console.log(curAgeRange + "s: " + femaleAgeGroups[i].length + " + " + maleAgeGroups[i].length + " (Female + Male)");
        curAgeRange += 10;
    }

    console.log("");

    console.log("Average of sharpness: " + sumOfSharpness / idx);

    console.log("");

    console.log(numbersOfSample + " samples are being extracted randomly for each age group in the order of sharpness.");

    // 남자 25명 + 여자 25명이 이상적이긴 함
    // 0. sharpness 순으로 정렬하기
    // 1. 남자 25개까지 채우기(addedMaleCount)
    // 2. addedFemaleCount = (50)-(addedMaleCount)
    
    curAgeRange = 30;
    for(let age = 0; age < 5; age++){
        
        if(maleAgeGroups[age].length >= 25){
            for(let maleCnt = 0; maleCnt < 25; maleCnt++){
                let curMaleID = maleAgeGroups[age][maleCnt];
                resultIDs[age].push(curMaleID);
            }
        }else{
            let numOfMale = maleAgeGroups[age].length;
            for(let maleCnt = 0; maleCnt < numOfMale; maleCnt++){
                let curMaleID = maleAgeGroups[age][maleCnt];
                resultIDs[age].push(curMaleID);
            }
        }

        let addedFemaleCount = numbersOfSample - resultIDs[age].length;
        if(femaleAgeGroups[age].length >= addedFemaleCount){
            for(let femaleCnt = 0; femaleCnt < addedFemaleCount; femaleCnt++){
                let curFemaleID = femaleAgeGroups[age][femaleCnt];
                resultIDs[age].push(curFemaleID);
            }
        }else{
            let numOfFemale = femaleAgeGroups[age].length;
            for(let femaleCnt = 0; femaleCnt < numOfFemale; femaleCnt++){
                let curFemaleID = femaleAgeGroups[age][femaleCnt];
                resultIDs[age].push(curFemaleID);
            }
        }
        
        //console.log(resultIDs[age].length);
        if(resultIDs[age].length < 50){
            console.log("");

            console.log("In " + curAgeRange + "s: " + resultIDs[age].length + " (The number of samples is less than " + numbersOfSample + ".)");
        }
        curAgeRange += 10;
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

    console.log("");
    console.log("");
    console.log("");
    console.log("");
    console.log("");
    sheetIdx++; // next sheet
}

// files 엑셀파일을 생성하고 저장한다.
//xlsx.writeFile(dataOutput, studyTerm + "_mined_data.xlsx"); 

