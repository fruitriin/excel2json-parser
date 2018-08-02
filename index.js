const xlsx = require('node-xlsx').default;
const fs=require("fs");

const workSheetsFromFile = xlsx.parse(`${__dirname}/WikiData.xlsm`);


const configs = {
    sheetLists :{
        "Spell" : {
            "targetColumns" : [
                "Name", "School", "Lv", "Main", "Sub", "Cost",
                "Use", "Terrain", "Type",
                "Damage", "Range", "Time", "AoE", "Number", "Prec", "Fatigue",
                "Spec", "ShortDesc", "Next"

            ]
        }
    }
};

const sheets = workSheetsFromFile.filter((sheet)=>{
    return Object.keys(configs.sheetLists).indexOf(sheet.name) !== -1
});

for(let sheet of sheets){
    const sheetname = sheet.name;
    let newdata = [];
    let index = new Array();
    for(let rowNum in sheet.data){
        row = sheet.data[rowNum];

        if(index.length === 0){
            for(let headNum in row){
                let head = row[headNum];
                let ordinal = configs.sheetLists[sheetname].targetColumns.indexOf(head)
                if(ordinal !== -1){
                    index[ordinal] = {"name" :head,  "index": headNum};
                }
            }
            continue;
        }

        let newrow = {};
        for(let col of index){
            newrow[col.name] = row[col.index] !== undefined ? row[col.index] : null;
        }
        newdata.push(newrow);
    }

    console.log(newdata)
    const resultJson = JSON.stringify(newdata,undefined,1)
    // console.log(resultJson);

    fs.writeFileSync( "./output/" + sheetname + ".json", resultJson);

}

