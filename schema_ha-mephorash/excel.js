/* **************************************************************************************************
 *
 * Name: Excel.js
 * Use: データ整形ツール
 * Description: Excel(xlsx)ファイルから `data.json` を作成
 *
 * Excel Format:
 *    - UTF-8/BOMなし
 *    - 改行コードLF
 *    - 空行なし
 *    - 各行の指定
 *        - col1: 名前
 *        - col2: カナ
 *        - col3: url
 *            - URL先が開けない場合は決め打ち文字列`404`とする
 *        - col4: 行
 *            - あ・か・さ……・わ、のいずれか
 *            - カタカナはひらがなに変換するのでカタカナひらがなどちらでもOK。ただし全角であること
 *        - col5: カテゴリ
 *            - 複数の場合、半角カンマ区切りで列挙
 * Using:
 *    1. 変換したいExcelファイルを `/src/` に保存
 *    2. `npm start <Excel Filename> <Sheet Name>` で本スクリプトを実行
 *    3. `/dist/` に `data.json` が生成される
 *
 ************************************************************************************************** */

//ライブラリ等読み込み
const fs        = require('fs')
const xlsx      = require('xlsx')
const utils     = xlsx.utils
const dir       = require('./dir')
const functions  = require('./functions')
const messages = require('./messages')
const argv = process.argv

//メイン処理
if(functions.fpCheck()) { //ファイルチェック成功ならば処理実行
    //引数チェック
    try {
        if(!functions.argvCheck(argv, functions)) {
            throw new Error(messages.argvLength)
        }
        else {
            //読み込み
            const dataExcel = book = xlsx.readFile(`${dir.el73.src}${argv[2]}`) //ファイル読み込み
            let arrayDist = []
            try {
                if(functions.mainLoop(dataExcel, arrayDist, argv, functions)) { //処理成功した場合
                    const dataJSON = JSON.stringify(arrayDist, undefined, 4) //2番目はフィルタ,3番目はインデント(半角スペース4文字)
                    //書き込み
                    fs.writeFileSync(dir.el73.dist, dataJSON)
                }
                else {
                    throw new Error(messages.processFailed)
                }
            } catch(e) {
                console.log(e)
            }
        }
    } catch(e) {
        console.log(e)
    }
}