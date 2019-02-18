const fs        = require('fs')
const xlsx      = require('xlsx')
const utils     = xlsx.utils
const dir       = require('./dir')
const messages  = require('./messages')

module.exports = {
    fpCheck: () => {
        try {
            //ファイル存在チェック
            fs.statSync(dir.el73.src)
            return true
        } catch(err) {
            if(err.code === 'ENOENT') console.log(messages.fpNotFound)
            return false
        }
    },
    argvCheck: (argv, functions) => {
        if(argv.length === 4) {
            for(let i = 0; i < argv.length; i++) {
                if(!functions.strCheck(argv[i])) return false
            }
            return true
        }
        else return false
    },
    filenameCheck: (str) => { //ファイル名チェック(/や\、..を削除)
        return str.replace(/(\/|\\|\.\.)+/g, '')
    },
    strCheck: (str) => { //文字列型かどうかチェック
        if(typeof str === 'string') return true
        else {
            console.log(`${str}${messages.notString}`)
            return false
        }
    },
    delimiterBreak: (str, delimiter) => { //デリミタでブレイク
        return str.split(delimiter)
    },
    lengthCheck: (str, i, j) => { //文字列の長さチェック
        if(str.length > 0) return true
        else {
            console.log(`${i}${messages.irregularColumn[0]}${j}${messages.irregularColumn[1]}`)
            return false
        }
    },
    kana2Hira: (str) => { //カタカナを平仮名に変換
        return str.replace(/[\u30a1-\u30f6]/g, function(match) {
            const chr = match.charCodeAt(0) - 0x60
            return String.fromCharCode(chr)
        })
    },
    targetLost: (url) => { //URLが404だった場合は空文字列、そうでない場合はURLをそのまま返す
        if(url === '404') {
            return ''
        }
        else {
            return url
        }
    },
    mainLoop: (dataExcel, arrayDist, argv, functions) => { //ループ
        const sheet = dataExcel.Sheets[argv[3]] //シート取得
        const range = sheet["!ref"] //範囲取得
        const decodeRange = utils.decode_range(range) //セル範囲を数値表現に変換
        for (let rowIndex = decodeRange.s.r; rowIndex <= decodeRange.e.r; rowIndex++) {
            let val = ['', '', '', '', []]
            for (let colIndex = decodeRange.s.c; colIndex <= decodeRange.e.c; colIndex++) {
                //データ列数チェック(5列でないデータはエラーで止まる)
                if(decodeRange.e.c !== 4) {
                    console.log(`${rowIndex}${messages.irregularLength}`)
                    return false
                }
                // 数値表現をセルアドレス ("A1"など) に変換
                const address = utils.encode_cell({ r: rowIndex, c:colIndex })
                const cell = sheet[address]
                if (typeof cell !== "undefined" && typeof cell.w !== "undefined") {
                    if(!functions.lengthCheck(cell.w, rowIndex, colIndex)) return false
                    else if(!functions.strCheck(cell.w)) return false
                    else val[colIndex] = cell.w
                }
            }
            const catArray = functions.delimiterBreak(val[4], ",") //カンマ区切り分解
            arrayDist.push({
                "name": val[0],
                "kana": val[1],
                "url": functions.targetLost(val[2]),
                "gyou": functions.kana2Hira(val[3]),
                "category": catArray
            })
        }
        return true
    }
}