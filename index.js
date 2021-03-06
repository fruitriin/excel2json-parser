const xlsx = require('node-xlsx').default;
const fs = require('fs');

const configs = JSON.parse(fs.readFileSync('./config.json', 'utf-8'))

const workSheetsFromFile = xlsx.parse(`${__dirname}/${configs.srcFile}`);
// config記載のシートのみに絞り込む
const sheets = workSheetsFromFile.filter((sheet) => {
    return Object.keys(configs.sheetLists).indexOf(sheet.name) !== -1
});

// シートデータを順繰りに処理
for (let sheet of sheets) {
    let sheetName = sheet.name;
    let indexes = {};
    let sheetConfig = configs.sheetLists[sheetName];

    let newData = sheetConfig.keyIndex === undefined ? [] : {};

    let skipKey = 0; //スキップ用のカラムが何番目か

    // データの各行にループ
    Object.keys(sheet.data).forEach(rowNum => {
        row = sheet.data[rowNum];

        // １行目だけカラム名が何行目にあるか採取
        if (Object.keys(indexes).length === 0) {
            // スキップ判定の列番号を採番
            skipKey = row.indexOf(sheetConfig.skipKey);
            // 出力用の列番号をキー名と列番号のペアで採番
            Object.keys(row).forEach(headNum => {
                let head = row[headNum];
                let ordinal = configs.sheetLists[sheetName].targetColumns.indexOf(head);
                if (ordinal !== -1) {
                    indexes[head] = headNum;
                }
            });
            return;
        }
        // スキップ判定に使うキーが 空の行は処理をスキップする
        if(row[skipKey] === undefined) return

        // 行を作成
        let newRow = {};
        // indexesから値データを詰めていく
        Object.keys(indexes).forEach(colName => {
            newRow[colName] = row[indexes[colName]] !== undefined ? row[indexes[colName]] : ''
        })

        if(sheetConfig.keyIndex === undefined){
            newData.push(newRow);
        }else{
            newData[row[indexes[sheetConfig.keyIndex]]] = newRow
        }
        
    });

    const resultJson = JSON.stringify(newData, undefined, 1)

    fs.mkdir('./output', () => {
    })
    fs.writeFileSync("./output/" + sheetName + ".json", resultJson);

}

