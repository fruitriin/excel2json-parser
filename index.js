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

            ],
            "skipKey" : "Name"
        },
        "NationalSpell" :{
            "targetColumns" : [
                "Name", "School", "Lv", "Main", "Sub", "Cost",
                "Use", "Terrain", "Type",
                "Damage", "Range", "Time", "AoE", "Number", "Prec", "Fatigue",
                "Spec", "ShortDesc", "Next"

            ],
            "skipKey" : "Name"
        },
        "Item": {
            "targetColumns" : [
                "Name","Pname","Main","Sub","Mc","Sc","Dam","DType","Arm","LD","Cap","Holy","AM","Size","Str","Cha","Re","Two","Fla","Head","Resi","Bon","NbrA","AoE","Lin","Att","Def","Ran","Pre","Amm","UW","Len","HP","BP","SP","Ade","Par","Enc","MMp","Mag","OnD","OnH","OnA","FR","SR","CR","PR","AR","IP","Inv","PF","Awe","AA","Fear","He","Ch","FS","AS","OC","PoB","Pet","DR","Bless","Ber","GB","Tra","DV","Reg","Rei","Ste","invis","Ass","Sed","Rea","Her","Inq","AdS","DoS","GoM","Ins","TM","BM","Co","Mco","Uco","Pill","PB","SupB","WB","Sail","Heal","Mas","Eth","Swi","Qui","Enl","StP","Luck","TF","FL","Curse","Taint","Dise","Eye","Che","Fee","Insa","Sie","Fly","SI","Flo","FSu","MSu","SSu","WSu","SnM","Swim","BS","ReB","DM","IL","ACB","ForB","MaS","MaR","Alc","CoM","CoS","Res","F","A","W","E","S","D","N","B","H","FC","GG","TG","Spec"
            ],
            "skipKey" : "Name"
        },
        "NationalItem": {
            "targetColumns" : [
                "Name","Pname","Main","Sub","Mc","Sc","Dam","DType","Arm","LD","Cap","Holy","AM","Size","Str","Cha","Re","Two","Fla","Head","Resi","Bon","NbrA","AoE","Lin","Att","Def","Ran","Pre","Amm","UW","Len","HP","BP","SP","Ade","Par","Enc","MMp","Mag","OnD","OnH","OnA","FR","SR","CR","PR","AR","IP","Inv","PF","Awe","AA","Fear","He","Ch","FS","AS","OC","PoB","Pet","DR","Bless","Ber","GB","Tra","DV","Reg","Rei","Ste","invis","Ass","Sed","Rea","Her","Inq","AdS","DoS","GoM","Ins","TM","BM","Co","Mco","Uco","Pill","PB","SupB","WB","Sail","Heal","Mas","Eth","Swi","Qui","Enl","StP","Luck","TF","FL","Curse","Taint","Dise","Eye","Che","Fee","Insa","Sie","Fly","SI","Flo","FSu","MSu","SSu","WSu","SnM","Swim","BS","ReB","DM","IL","ACB","ForB","MaS","MaR","Alc","CoM","CoS","Res","F","A","W","E","S","D","N","B","H","FC","GG","TG","Spec"
            ],
            "skipKey" : "Name"
        },


    }
};

const sheets = workSheetsFromFile.filter((sheet)=>{
    return Object.keys(configs.sheetLists).indexOf(sheet.name) !== -1
});

for(let sheet of sheets){
    let sheetname = sheet.name;
    let newdata = [];
    let index = new Array();
    for(let rowNum in sheet.data){
        row = sheet.data[rowNum];
        let skipKey = 0;
        if(index.length === 0){
            skipKey = row.indexOf(configs.sheetLists[sheetname].skipKey);
            for(let headNum in row){
                let head = row[headNum];
                let ordinal = configs.sheetLists[sheetname].targetColumns.indexOf(head);
                if(ordinal !== -1){
                    index[ordinal] = {"name" :head,  "index": headNum};
                }
            }
            continue;
        }


        if(row[skipKey] === undefined) continue;
        let newrow = {};
        for(let col of index){
            newrow[col.name] = row[col.index] !== undefined ? row[col.index] : null;
        }
        newdata.push(newrow);
    }

    // console.log(newdata)
    const resultJson = JSON.stringify(newdata,undefined,1)
    // console.log(resultJson);

    fs.writeFileSync( "./output/" + sheetname + ".json", resultJson);
    fs.mkdir('./output', () => {
    })

}

