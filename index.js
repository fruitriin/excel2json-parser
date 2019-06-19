const xlsx = require('node-xlsx').default;
const fs = require('fs');

const workSheetsFromFile = xlsx.parse(`${__dirname}/WikiData.xlsm`);

const configs = {
    sheetLists: {
        "Spell": {
            "targetColumns": [
                "Name", "School", "Lv", "Main", "Sub", "Cost",
                "Use", "Terrain", "Type",
                "Damage", "Range", "Time", "AoE", "Number", "Prec", "Fatigue",
                "Spec", "ShortDesc", "Next"

            ],
            "skipKey": "Name"
        },
        "NationalSpell": {
            "targetColumns": [
                "Name", "School", "Lv", "Main", "Sub", "Cost",
                "Use", "Terrain", "Type",
                "Damage", "Range", "Time", "AoE", "Number", "Prec", "Fatigue",
                "Spec", "ShortDesc", "Next"

            ],
            "skipKey": "Name"
        },
        "Item": {
            "targetColumns": [
                "Name", "Pname", "Main", "Sub", "Mc", "Sc", "Dam", "DType", "Arm", "LD", "Cap", "Holy", "AM", "Size", "Str", "Cha", "Re", "Two", "Fla", "Head", "Resi", "Bon", "NbrA", "AoE", "Lin", "Att", "Def", "Ran", "Pre", "Amm", "UW", "Len", "HP", "BP", "SP", "Ade", "Par", "Enc", "MMp", "Mag", "OnD", "OnH", "OnA", "FR", "SR", "CR", "PR", "AR", "IP", "Inv", "PF", "Awe", "AA", "Fear", "He", "Ch", "FS", "AS", "OC", "PoB", "Pet", "DR", "Bless", "Ber", "GB", "Tra", "DV", "Reg", "Rei", "Ste", "invis", "Ass", "Sed", "Rea", "Her", "Inq", "AdS", "DoS", "GoM", "Ins", "TM", "BM", "Co", "Mco", "Uco", "Pill", "PB", "SupB", "WB", "Sail", "Heal", "Mas", "Eth", "Swi", "Qui", "Enl", "StP", "Luck", "TF", "FL", "Curse", "Taint", "Dise", "Eye", "Che", "Fee", "Insa", "Sie", "Fly", "SI", "Flo", "FSu", "MSu", "SSu", "WSu", "SnM", "Swim", "BS", "ReB", "DM", "IL", "ACB", "ForB", "MaS", "MaR", "Alc", "CoM", "CoS", "Res", "F", "A", "W", "E", "S", "D", "N", "B", "H", "FC", "GG", "TG", "Spec"
            ],
            "skipKey": "Name"
        },
        "NationalItem": {
            "targetColumns": [
                "Name", "Pname", "Main", "Sub", "Mc", "Sc", "Dam", "DType", "Arm", "LD", "Cap", "Holy", "AM", "Size", "Str", "Cha", "Re", "Two", "Fla", "Head", "Resi", "Bon", "NbrA", "AoE", "Lin", "Att", "Def", "Ran", "Pre", "Amm", "UW", "Len", "HP", "BP", "SP", "Ade", "Par", "Enc", "MMp", "Mag", "OnD", "OnH", "OnA", "FR", "SR", "CR", "PR", "AR", "IP", "Inv", "PF", "Awe", "AA", "Fear", "He", "Ch", "FS", "AS", "OC", "PoB", "Pet", "DR", "Bless", "Ber", "GB", "Tra", "DV", "Reg", "Rei", "Ste", "invis", "Ass", "Sed", "Rea", "Her", "Inq", "AdS", "DoS", "GoM", "Ins", "TM", "BM", "Co", "Mco", "Uco", "Pill", "PB", "SupB", "WB", "Sail", "Heal", "Mas", "Eth", "Swi", "Qui", "Enl", "StP", "Luck", "TF", "FL", "Curse", "Taint", "Dise", "Eye", "Che", "Fee", "Insa", "Sie", "Fly", "SI", "Flo", "FSu", "MSu", "SSu", "WSu", "SnM", "Swim", "BS", "ReB", "DM", "IL", "ACB", "ForB", "MaS", "MaR", "Alc", "CoM", "CoS", "Res", "F", "A", "W", "E", "S", "D", "N", "B", "H", "FC", "GG", "TG", "Spec"
            ],
            "skipKey": "Name"
        },


    }
};


// config記載のシートのみに絞り込む
const sheets = workSheetsFromFile.filter((sheet) => {
    return Object.keys(configs.sheetLists).indexOf(sheet.name) !== -1
});

// シートデータを順繰りに処理
for (let sheet of sheets) {
    let sheetName = sheet.name;
    let newData = [];
    let indexes = [];

    // データの各行にループ
    Object.keys(sheet.data).forEach(rowNum => {
        row = sheet.data[rowNum];
        let skipKey = 0; // スキップ判定用のカラムが何番目の列か

        // １行目だけカラム名が何行目にあるか採取
        if (indexes.length === 0) {
            // スキップ判定の列番号を採番
            skipKey = row.indexOf(configs.sheetLists[sheetName].skipKey);

            // 出力用の列番号をキー名と列番号のペアで採番
            Object.keys(row).forEach(headNum => {
                let head = row[headNum];
                let ordinal = configs.sheetLists[sheetName].targetColumns.indexOf(head);
                if (ordinal !== -1) {
                    indexes[ordinal] = {"name": head, "skipKey": headNum};
                }
            });
            return;
        }

        // 行を作成
        let newRow = {};
        // indexesから値データを詰めていく
        for (let col of indexes) {
            newRow[col.name] = row[col.skipKey] !== undefined ? row[col.skipKey] : '';
        }

        newData.push(newRow);
    });

    const resultJson = JSON.stringify(newData, undefined, 1)

    fs.mkdir('./output', () => {
    })
    fs.writeFileSync("./output/" + sheetName + ".json", resultJson);

}

