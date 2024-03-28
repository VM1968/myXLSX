let MXLSX = {};
MXLSX.SelectedRow=null;
//FileReader async 
class FileReaderEx extends FileReader {
    constructor() {
        super();
    }

    readAs(blob, ctx) {
        return new Promise((res, rej) => {
            super.addEventListener("load", ({ target }) => res(target.result));
            super.addEventListener("error", ({ target }) => rej(target.error));
            super[ctx](blob);
        });
    }
    readAsEnc(blob, encoding, ctx) {
        return new Promise((res, rej) => {
            super.addEventListener("load", ({ target }) => res(target.result));
            super.addEventListener("error", ({ target }) => rej(target.error));
            super[ctx](blob, encoding);
        });
    }

    readAsArrayBuffer(blob) {
        return this.readAs(blob, "readAsArrayBuffer");
    }

    readAsDataURL(blob) {
        return this.readAs(blob, "readAsDataURL");
    }

    readAsText(blob) {
        return this.readAs(blob, "readAsText");
    }

    readAsTextEnc(blob, encoding) {
        console.log(encoding);
        return this.readAsEnc(blob, encoding, "readAsText");
    }
}

const files = {};

const position = ["left", "right", "top", "bottom"];


//кодировки цвета 
https://stackoverflow.com/questions/2353211/hsl-to-rgb-color-conversion
function RGBToHSL(R, G, B) {
    let r = R / 255;
    let g = G / 255;
    let b = B / 255;
    min = Math.min(r, Math.min(g, b));
    max = Math.max(r, Math.max(g, b));
    delta = max - min;
    if (max == min) {
        H = 0;
        S = 0;
        L = max;
        return [H, S * 100, L * 100];
    }
    L = (min + max) / 2;
    if (L < 0.5) {
        S = delta / (max + min);
    }
    else {
        S = delta / (2.0 - max - min);
    }
    if (r == max) H = (g - b) / delta;
    if (g == max) H = 2.0 + (b - r) / delta;
    if (b == max) H = 4.0 + (r - g) / delta;
    H *= 60;
    if (H < 0) H += 360;
    // return [Math.round(H), Math.round(S * 100), Math.round(L * 100)];
    return [H, S * 100, L * 100];
}
function HSLToRGB(h, s, l) {
    s /= 100;
    l /= 100;
    const k = n => (n + h / 30) % 12;
    const a = s * Math.min(l, 1 - l);
    const f = n =>
        l - a * Math.max(-1, Math.min(k(n) - 3, Math.min(9 - k(n), 1)));
    return [Math.round(255 * f(0)), Math.round(255 * f(8)), Math.round(255 * f(4))];
};
function rgbToHex(r, g, b) {
    let rhex = r.toString(16);
    let ghex = g.toString(16);
    let bhex = b.toString(16);
    return (rhex.length == 1 ? "0" + rhex : rhex) + (ghex.length == 1 ? "0" + ghex : ghex) + (bhex.length == 1 ? "0" + bhex : bhex);
}
function CalculateFinalLumValue(tint, lum) {
    if (!tint) {
        return lum;
    }
    lum1 = 0;
    if (Number(tint) < 0) {
        lum1 = lum * (1.0 + Number(tint));
    }
    else {
        lum1 = lum * (1.0 + Number(tint));// + (255 - 255 * (1.0 - Number(tint)));
    }
    return lum1;
}
function indexedToRGB(h, tint) {
    //console.log(tint);
    //console.log(h);
    let aRgbHex = h.match(/^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i);
    r = parseInt(aRgbHex[1], 16);
    g = parseInt(aRgbHex[2], 16);
    b = parseInt(aRgbHex[3], 16);

    //return [r, g, b];
    //RGBToHSL
    HSL = RGBToHSL(r, g, b);
    //console.log(HSL);
    let nL = CalculateFinalLumValue(tint, HSL[2]);
    //console.log();
    nL = Math.round(nL, 0);
    //console.log(nL);
    RGB = HSLToRGB(HSL[0], HSL[1], nL);
    //return RGB;
    hex = rgbToHex(RGB[0], RGB[1], RGB[2]);
    //console.log(hex);
    return hex;
}


// проверка подходит ли для отображения дат
function is_datefmt(fmt) {
    const regex = /[msyh]/g;
    if ((m = regex.exec(fmt)) !== null) { return true }
    else { return false }
}
//В Excel свой способ хранения дат
function parseExcelDate(excelSerialDate) {
    //(Excel_time-DATE(1970,1,1))*86400 с   //*1000 мс
    return new Date(Math.round((excelSerialDate - 25569) * 86400000));
};


function findMerged(idcell) {
    //находится ли в списке объединяемых ячеек с дополнительными параметрами "rowspan","colspan"
    //иначе обычная видимая ячейка
    let fcell = {};

    fcell = mcells.find(el => el.id === idcell);

    if (!fcell) {
        fcell = { "id": idcell, "enabled": true }
    }
    return fcell;
}


//номер колонки по ID ячейки
function idc(str) {
    //id column for cell "A5" = 1
    const regex = /[A-Z]{1,}/;
    let m;
    let id = 0;
    m = regex.exec(str);
    //console.log(m[0]);
    for (let i = 0; i < m[0].length; i++) {
        id += m[0].charCodeAt(i) % 64;
    }
    return id
}

//номер строки по ID ячейки
function idr(str) {
    //id row for cell "A5" = 5
    const regex = /[0-9]{1,}/;
    let m;
    let id = 0;
    m = regex.exec(str);
    id = Number(m[0]);
    return id
}

//https://gist.github.com/chinchang/8106a82c56ad007e27b1
function xmlToJson(xml) {
    let js_obj = {};
    if (xml.nodeType == 1) {
        if (xml.attributes.length > 0) {
            js_obj["@attributes"] = {};
            for (let j = 0; j < xml.attributes.length; j++) {
                let attribute = xml.attributes.item(j);
                js_obj["@attributes"][attribute.nodeName] = attribute.value;
            }
        }
    } else if (xml.nodeType == 3) {
        js_obj = xml.nodeValue;
    }
    if (xml.hasChildNodes()) {
        for (let i = 0; i < xml.childNodes.length; i++) {
            let item = xml.childNodes.item(i);
            let nodeName = item.nodeName;
            if (typeof (js_obj[nodeName]) == "undefined") {
                js_obj[nodeName] = xmlToJson(item);
            } else {
                if (typeof (js_obj[nodeName].push) == "undefined") {
                    let old = js_obj[nodeName];
                    js_obj[nodeName] = [];
                    js_obj[nodeName].push(old);
                }
                js_obj[nodeName].push(xmlToJson(item));
            }
        }
    }
    return js_obj;
}

//numFmtId
function init_table(t) {
    t[0] = 'General';
    t[1] = '0';
    t[2] = '0.00';
    t[3] = '#,##0';
    t[4] = '#,##0.00';
    t[9] = '0%';
    t[10] = '0.00%';
    t[11] = '0.00E+00';
    t[12] = '# ?/?';
    t[13] = '# ??/??';
    t[14] = 'dd/mm/yyyy';
    t[15] = 'd-mmm-yy';
    t[16] = 'd-mmm';
    t[17] = 'mmm-yy';
    t[18] = 'h:mm AM/PM';
    t[19] = 'h:mm:ss AM/PM';
    t[20] = 'h:mm';
    t[21] = 'h:mm:ss';
    t[22] = 'd/mm/yy h:mm';
    t[37] = '#,##0 ;(#,##0)';
    t[38] = '#,##0 ;[Red](#,##0)';
    t[39] = '#,##0.00;(#,##0.00)';
    t[40] = '#,##0.00;[Red](#,##0.00)';
    t[45] = 'mm:ss';
    t[46] = '[h]:mm:ss';
    t[47] = 'mmss.0';
    t[48] = '##0.0E+0';
    t[49] = '@';
    t[56] = '"上午/下午 "hh"時"mm"分"ss"秒 "';
}

// //пока не использовал
// let charset = [];
// charset[0] = "Ansi"
// charset[1] = "Default"
// charset[2] = "Symbol"
// charset[77] = "Mac"
// charset[128] = "ShiftJIS"
// charset[129] = "Hangeul"
// charset[130] = "Johab"
// charset[134] = "GB2312"
// charset[136] = "ChineseBig5"
// charset[161] = "Greek"
// charset[162] = "Turkish"
// charset[163] = "Vietnamese"
// charset[177] = "Hebrew"
// charset[178] = "Arabic"
// charset[186] = "Baltic"
// charset[204] = "Russian"
// charset[222] = "Thai"
// charset[238] = "EastEurope"
// charset[255] = "Oem"

//indexed color c# ARGB Value
//ID заданных стандартных  цветов
let indexed = [];
indexed[0] = "00000000";
indexed[1] = "00FFFFFF";
indexed[2] = "00FF0000";
indexed[3] = "0000FF00";
indexed[4] = "000000FF";
indexed[5] = "00FFFF00";
indexed[6] = "00FF00FF";
indexed[7] = "0000FFFF";
indexed[8] = "00000000";
indexed[9] = "00FFFFFF";
indexed[10] = "00FF0000";
indexed[11] = "0000FF00";
indexed[12] = "000000FF";
indexed[13] = "00FFFF00";
indexed[14] = "00FF00FF";
indexed[15] = "0000FFFF";
indexed[16] = "00800000";
indexed[17] = "00008000";
indexed[18] = "00000080";
indexed[19] = "00808000";
indexed[20] = "00800080";
indexed[21] = "00008080";
indexed[22] = "00C0C0C0";
indexed[23] = "00808080";
indexed[24] = "009999FF";
indexed[25] = "00993366";
indexed[26] = "00FFFFCC";
indexed[27] = "00CCFFFF";
indexed[28] = "00660066";
indexed[29] = "00FF8080";
indexed[30] = "000066CC";
indexed[31] = "00CCCCFF";
indexed[32] = "00000080";
indexed[33] = "00FF00FF";
indexed[34] = "00FFFF00";
indexed[35] = "0000FFFF";
indexed[36] = "00800080";
indexed[37] = "00800000";
indexed[38] = "00008080";
indexed[39] = "000000FF";
indexed[40] = "0000CCFF";
indexed[41] = "00CCFFFF";
indexed[42] = "00CCFFCC";
indexed[43] = "00FFFF99";
indexed[44] = "0099CCFF";
indexed[45] = "00FF99CC";
indexed[46] = "00CC99FF";
indexed[47] = "00FFCC99";
indexed[48] = "003366FF";
indexed[49] = "0033CCCC";
indexed[50] = "0099CC00";
indexed[51] = "00FFCC00";
indexed[52] = "00FF9900";
indexed[53] = "00FF6600";
indexed[54] = "00666699";
indexed[55] = "00969696";
indexed[56] = "00003366";
indexed[57] = "00339966";
indexed[58] = "00003300";
indexed[59] = "00333300";
indexed[60] = "00993300";
indexed[61] = "00993366";
indexed[62] = "00333399";
indexed[63] = "00333333";
indexed[64] = "00000000";//"SystemForeground"; черный
indexed[65] = "00FFFFFF";//"SystemBackground"; белый

//ID заданных стандартных  цветовых схем
let themecolor = [];
themecolor[0] = "ffffff";
themecolor[1] = "000000";
themecolor[2] = "e7e6e6";
themecolor[3] = "44546a";
themecolor[4] = "4472c4";
themecolor[5] = "ed7d31";
themecolor[6] = "a5a5a5";
themecolor[7] = "ffc000";
themecolor[8] = "5b9bd5";
themecolor[9] = "70ad47";


//дополнительные пользовательские форматы в Excel
function parseStyle(xlStyle) {
    //console.log(xlStyle);
    //No CSS rule
    let styles = {};

    //NumberFmt
    //Числовые форматы
    let numFmt = [];

    //стандартные числовые форматы
    init_table(numFmt);
    //дополнительные форматы для чисел, заданные в файле EXCEL
    if ("numFmts" in xlStyle.styleSheet) {
        if (Array.isArray(xlStyle.styleSheet.numFmts.numFmt)) {
            xlStyle.styleSheet.numFmts.numFmt.forEach(fmt => {
                numFmt[fmt["@attributes"]["numFmtId"]] = fmt["@attributes"]["formatCode"]
            })
        } else {
            numFmt[xlStyle.styleSheet.numFmts.numFmt["@attributes"]["numFmtId"]] = xlStyle.styleSheet.numFmts.numFmt["@attributes"]["formatCode"]
        }

    }
    styles.NumberFmt = numFmt

    //CellXf
    let cellxfs = [];
    if (Array.isArray(xlStyle.styleSheet.cellXfs.xf)) {
        xlStyle.styleSheet.cellXfs.xf.forEach(cxf => {
            cellxf = {};
            cellxf.numFmt = numFmt[cxf["@attributes"]["numFmtId"]];
            cellxfs.push(cellxf);
        })
    } else {

    }

    styles.CellXf = cellxfs;
    return styles;
}

function parseStyleCSS(xlStyle) {
    //CSS rule
    //Fonts CSS
    let fontsCSS = [];

    function readFont(fnt) {
        //console.log(fnt);
        let fontcss = '';
        if (("name" in fnt)) {
            fontcss += 'font-family: ' + fnt.name["@attributes"]["val"] + ';';
        }
        if (("sz" in fnt)) {
            fontcss += 'font-size: ' + Number(fnt.sz["@attributes"]["val"]) + 'pt;';
        }
        //color: #FFFF0000;
        //ARGB 
        if ("color" in fnt) {
            if ("theme" in fnt.color["@attributes"]) {
                if (themecolor[fnt.color["@attributes"]["theme"]]) {
                    fontcss += 'color: #' + themecolor[fnt.color["@attributes"]["theme"]] + ';';
                }
            }
            if ("indexed" in fnt.color["@attributes"]) {
                if (indexed[fnt.color["@attributes"]["indexed"]]) {
                    fontcss += 'color: #' + indexed[fnt.color["@attributes"]["indexed"]].slice(2, 8) + ';';
                }
            }
            if ("rgb" in fnt.color["@attributes"]) {
                fontcss += 'color: #' + fnt.color["@attributes"]["rgb"].slice(2, 8) + ';';
            }
        };
        //font-weight
        if ("b" in fnt) {
            fontcss += 'font-weight: bold;'
        };
        //font-style
        if ("i" in fnt) {
            fontcss += 'font-style: italic;'
        };
        //text-decoration
        if ("u" in fnt) {
            fontcss += 'text-decoration: underline;'
        };

        fontsCSS.push(fontcss);
    }

    if (Array.isArray(xlStyle.styleSheet.fonts.font)) {
        xlStyle.styleSheet.fonts.font.forEach(fnt => {
            readFont(fnt);
        })
    } else {
        fnt = xlStyle.styleSheet.fonts.font;
        readFont(fnt);
    }
    // console.log(fontsCSS);

    //Fill CSS
    let fillsCSS = [];
    if (Array.isArray(xlStyle.styleSheet.cellXfs.xf)) {
        xlStyle.styleSheet.fills.fill.forEach(fl => {
            fillcss = '';

            if ("patternFill" in fl) {
                if ("fgColor" in fl.patternFill) {
                    if ("rgb" in fl.patternFill.fgColor["@attributes"]) {
                        fillcss += 'background-color: #' + fl.patternFill.fgColor["@attributes"]["rgb"].slice(2, 8) + ' !important;';

                    }
                    if ("indexed" in fl.patternFill.fgColor["@attributes"]) {
                        fillcss += 'background-color: #' + indexed[fl.patternFill.fgColor["@attributes"]["indexed"]].slice(2, 8) + ' !important;';
                    }
                    if ("theme" in fl.patternFill.fgColor["@attributes"]) {
                        fillcss += 'background-color: #' + indexedToRGB(themecolor[fl.patternFill.fgColor["@attributes"]["theme"]], fl.patternFill.fgColor["@attributes"]["tint"]) + ' !important;';
                    }
                }
            };


            fillsCSS.push(fillcss);
        })
    }
    //    console.log(fillsCSS);

    //Border CSS
    let bordersCSS = [];

    if (Array.isArray(xlStyle.styleSheet.borders.border)) {
        xlStyle.styleSheet.borders.border.forEach(brd => {
            bordercss = '';

            position.forEach(ps => {
                if (ps in brd) {
                    clr = '';
                    if ("color" in brd[ps]) {
                        if ("indexed" in brd[ps]["color"]["@attributes"]) {
                            clr = ' #' + indexed[brd[ps]["color"]["@attributes"]["indexed"]].slice(2, 8);
                        }
                        if ("rgb" in brd[ps]["color"]["@attributes"]) {
                            clr = ' #' + brd[ps]["color"]["@attributes"]["rgb"].slice(2, 8);
                        }
                    }

                    if (["@attributes"] in brd[ps] && "style" in brd[ps]["@attributes"]) {
                        bordercss += 'border-' + ps + ': solid ' + brd[ps]["@attributes"]["style"] + clr + ' !important;'
                    }

                }
            })

            bordersCSS.push(bordercss);
        })
    } else {
        // на нет и суда нет
    }
    //   console.log(bordersCSS);

    //cellXfs CSS
    //let classCSS = [];

    //Было
    // classCSS = '<style>';
    // classCSS += 'table{border-collapse: collapse;table-layout: fixed;font-family: Calibri;font-size: 11pt;}';
    // classCSS += 'td{padding-top:2px;padding-right:2px;padding-left:5px;vertical-align: bottom;word-break: break-all;white-space: nowrap;}' //white-space: nowrap;
    // classCSS += 'td.Number {text-align: end;}';
    // classCSS += 'tr{height: 20px;}';
    // classCSS += 'tr:hover{background-color: #e6e6e6;}';
    // classCSS += '.inlineStr {white-space: nowrap;overflow: hidden;text-overflow: ellipsis;}'; //background-color: red;
    // classCSS += '.inlineStr:hover {white-space: normal;z-index: 1}';
    //Было

    //Стало
    classCSS = '<style>';
    classCSS += '.sheetjs{border-spacing: initial;table-layout: fixed;font-family: Calibri;font-size: 11pt;}';//border-spacing: 0.1px;border-collapse: collapse;
    classCSS += '.sheetjs td{border: 0.1px solid #f0f0f0;padding-top:1px;padding-right:1px;padding-left:1px;vertical-align: bottom; user-select: none;}'//  white-space: nowrap;}'   //word-break: break-all; border: 1px solid #ccc; //border-top: 1px solid #ccc;border-right: 1px solid #ccc;border-left: 1px solid #ccc;border-bottom: 1px solid #ccc;
    classCSS += '.sheetjs td.Number {text-align: end;word-break: unset;}';
    classCSS += '.sheetjs tr{height: 20px;}';
    classCSS += '.sheetjs tr:hover{background-color: #e6e6e6;}';
    classCSS += '.sheetjs .inlineStr {white-space: nowrap;overflow: hidden;text-overflow: ellipsis;}'; //background-color: red;
    classCSS += '.sheetjs .inlineStr:hover {white-space: normal;z-index: 1}';
    classCSS += '.btn {cursor: col-resize;/*background-color: red;*/user-select: none;position: absolute;}';
    classCSS += '.btn-right { right: 0; }';

    classCSS += '.sheetjs thead > tr > th {text-align: center; border-top: 1px solid #ccc;border-left: 1px solid #ccc;border-right: 1px solid #ccc;' +
        'border-bottom: 1px solid #ccc;background-color: #f3f3f3;cursor: pointer;box-sizing: border-box;' +
        'overflow: hidden; position: -webkit-sticky; position: sticky;top: 0;z-index: 2;}';//padding: 2px;

    classCSS += ".sheetjs thead > tr > th:hover {background-color: #a4ce88;}";

    classCSS += '.sheetjs .row_selectall { text-align: center; border-top: 1px solid #ccc;border-left: 1px solid #ccc;border-right: 1px solid #ccc;' +
        'border-bottom: 1px solid #ccc;background-color: #f3f3f3;cursor: pointer;box-sizing: border-box;' +
        'overflow: hidden; position: -webkit-sticky; position: sticky;top: 0;z-index: 2;word-break: unset;}';//padding: 2px;

    classCSS += ".sheetjs .row_selectall:hover {background-color: #a4ce88;}";

    //Стало

    if (Array.isArray(xlStyle.styleSheet.cellXfs.xf)) {
    xlStyle.styleSheet.cellXfs.xf.forEach(function (item, ind) {
        strstyle = '.cstyle' + ind + ' {';
        strstyle += fontsCSS[item["@attributes"]["fontId"]];
        strstyle += fillsCSS[item["@attributes"]["fillId"]];
        strstyle += bordersCSS[item["@attributes"]["borderId"]];

        if ('alignment' in item && 'horizontal' in item.alignment["@attributes"]) { strstyle += ' text-align:' + item.alignment["@attributes"].horizontal + ' !important;' }
        if ('alignment' in item && 'vertical' in item.alignment["@attributes"]) {
            if (item.alignment["@attributes"].vertical == 'center') { strstyle += ' vertical-align: middle !important;' }
            else { strstyle += ' vertical-align:' + item.alignment["@attributes"].vertical + ' !important;' }
        }

        if ('alignment' in item && 'wrapText' in item.alignment["@attributes"]) { strstyle += ' white-space: normal;' }
        else { strstyle += ' white-space: pre;' }

        strstyle += '}';
        //classCSS.push(strstyle);
        classCSS += strstyle;
    })
    }
    //console.log(classCSS);
    classCSS += '</style>';
    return classCSS;
}

let maxcellnumber = 0;
function parseSheet(XLsheet, sst, styles) {
    maxcellnumber = 0;
    let xlsheet = XLsheet;
    let rows = [];
    let mcolumn = {};

    function readCell(column) {
        //console.log(column);
        switch (column["@attributes"]["t"]) {
            case "s":
                //строка общая в отдельном массиве
                //console.log(sst[column.v["#text"]]);
                mcolumn = {
                    "id": column["@attributes"]["r"]
                };
                //console.log(sst[column.v["#text"]]);
                if ("r" in sst[column.v["#text"]]) {
                    let txt = '';
                    sst[column.v["#text"]]["r"].forEach(r => {
                        let stylePr = '';
                        if ("rPr" in r) {
                            stylePr = ` style="${readrPr(r.rPr)}"`;
                        }
                        txt += `<span ${stylePr}>${r["t"]["#text"]}</span>`;
                    })
                    mcolumn.v = txt;
                } else {
                    mcolumn.v = sst[column.v["#text"]]["t"]["#text"];
                    if ("@attributes" in sst[column.v["#text"]]["t"] && "xml:space" in sst[column.v["#text"]]["t"]["@attributes"]) {
                        mcolumn.whitespace = 'pre';
                    }
                }


                if ("s" in column["@attributes"]) {
                    mcolumn.style = column["@attributes"]["s"]//styles.CellXf[column["@attributes"]["s"]]
                }

                break;
            case "inlineStr":
                //строка непосредственно в ячейке
                mcolumn = {
                    "id": column["@attributes"]["r"],
                    "v": column.is.t["#text"],
                    "inlineStr": 1  //возможен длинный текст 
                };
                if ("s" in column["@attributes"]) {
                    mcolumn.style = column["@attributes"]["s"]//styles.CellXf[column["@attributes"]["s"]]
                }
                break;
            default:
                //числовое значение
                if (styles.CellXf[column["@attributes"]["s"]] && "numFmt" in styles.CellXf[column["@attributes"]["s"]]) {
                    fmt = styles.CellXf[column["@attributes"]["s"]]["numFmt"];
                } else { fmt = "" }
                if (fmt != "" && fmt != "General") {
                    //console.log(fmt);
                    v = "";
                    if (column.v) {
                        if (is_datefmt(fmt)) {
                            v = SSF.format(fmt, parseExcelDate(column.v["#text"]))
                            //v = SSF.format(fmt, column.v["#text"]) 
                        } else {
                            v = SSF.format(fmt, Number(column.v["#text"]))
                        }
                        v = v.replace(/\,/g, " ");//разделитель разрядов
                        v = v.replace(/\//g, ".");//разделитель разрядов
                    }
                } else {
                    if (column.v && "#text" in column.v) {
                        v = column.v["#text"]
                    } else { v = "" };
                }
                mcolumn = {
                    "id": column["@attributes"]["r"],
                    "v": v,
                    "type": "Number"
                };
                if ("s" in column["@attributes"]) {
                    mcolumn.style = column["@attributes"]["s"]//styles.CellXf[column["@attributes"]["s"]]
                }
                break;
        }
        c.push(mcolumn);
    }

    //читать ряд
    function readRow(element) {
        row = {};
        c = [];

        row.r = element["@attributes"]["r"];
        row.spans = element["@attributes"]["spans"];
        if ("ht" in element["@attributes"]) {
            row.ht = Math.round(element["@attributes"]["ht"]);
        }

        if (Array.isArray(element.c)) {
            //если массив столбцов
            element.c.forEach(celm => {
                readCell(celm);
            });
        } else {
            //одинокий столбец в строке ....если есть (может быть только параметры строки)
            if ("c" in element) {
                celm = element.c;
                readCell(celm);
            }
        }

        //console.log( c[c.length-1]);
        if (c.length > 0 && idc(c[c.length - 1].id) > maxcellnumber) { maxcellnumber = idc(c[c.length - 1].id) };


        row.columns = c;
        rows.push(row);
    }

    function readrPr(rPr) {
        fontcss = '';
        if (("rFont" in rPr)) {
            fontcss += 'font-family: ' + rPr.rFont["@attributes"]["val"] + ';';
        }
        if (("sz" in rPr)) {
            fontcss += 'font-size: ' + Number(rPr.sz["@attributes"]["val"]) + 'pt;';
        }
        //color: #FFFF0000;
        //ARGB 
        if ("color" in rPr) {
            if ("theme" in rPr.color["@attributes"]) {
                if (themecolor[rPr.color["@attributes"]["theme"]]) {
                    fontcss += 'color: #' + themecolor[rPr.color["@attributes"]["theme"]] + ';';
                }
            }
            if ("indexed" in rPr.color["@attributes"]) {
                if (indexed[rPr.color["@attributes"]["indexed"]]) {
                    fontcss += 'color: #' + indexed[rPr.color["@attributes"]["indexed"]].slice(2, 8) + ';';
                }
            }
            if ("rgb" in rPr.color["@attributes"]) {
                fontcss += 'color: #' + rPr.color["@attributes"]["rgb"].slice(2, 8) + ';';
            }
        };
        //font-weight
        if ("b" in rPr) {
            fontcss += 'font-weight: bold;'
        };
        //font-style
        if ("i" in rPr) {
            fontcss += 'font-style: italic;'
        };
        //text-decoration
        if ("u" in rPr) {
            fontcss += 'text-decoration: underline;'
        };
        return fontcss;
    }

    if (Array.isArray(xlsheet.row)) {
        xlsheet.row.forEach(elm => {
            readRow(elm);
        });
    } else {
        elm = xlsheet.row;
        readRow(elm);
    }
    return rows;
}

//описание <cols></cols>
function parseCols(Cols, maxcell) {
    //console.log(maxcell);
    let xlcols = {};
    let cols = [];

    let wd = 8.43;// ширина колонки по умолчанию
    //console.log(wd);
    if (Array.isArray(Cols.col)) {
        Cols.col.forEach(cl => {
            //если не описаны колонки до "min"
            min = parseInt(cl["@attributes"]["min"]); max = parseInt(cl["@attributes"]["max"]);

            for (let index = min; (cols.length < maxcell && index <= max); index++) {
                col = {};
                col.width = cl["@attributes"]["width"];
                //    width = col.width;
                cols[index - 1] = col;
            }
        })
    }
    else {
        //авто ширина колонок?
        //если не описаны колонки до "min"
        for (let i = 1; i < Number(Cols.col["@attributes"]["min"]); i++) {
            let col = {};
            col.width = wd;
            //    width = col.width;
            cols[i - 1] = col;
        }
        for (let i = Number(Cols.col["@attributes"]["min"]); cols.length < maxcell && i <= Number(Cols.col["@attributes"]["max"]); i++) {
            col = {};
            col.width = Cols.col["@attributes"]["width"];
            //    width = col.width;
            cols[i - 1] = col;
        }
    }
    xlcols.cols = cols;
    //xlcols.width = width;
    return xlcols;
}

function parseWB(xlWB) {

    //console.log(xlWB);
    let wb = {};
    wb.WBProps = { "date1904": false };

    //Sheets
    let sheets = [];

    if (Array.isArray(xlWB.workbook.sheets.sheet)) {
        xlWB.workbook.sheets.sheet.forEach(sht => {
            sheet = {};
            sheet.name = sht["@attributes"]["name"];
            sheet.sheetId = sht["@attributes"]["sheetId"];
            sheet.file = sht["@attributes"]["r:id"].replace("rId", "sheet") + '.xml';
            if ("state" in sht["@attributes"] && sht["@attributes"].state == "hidden") {
                sheet.hidden = true;
            }
            sheets.push(sheet);
        })
    } else {
        sheet = {};
        sheet.name = xlWB.workbook.sheets.sheet["@attributes"]["name"];
        sheet.sheetId = xlWB.workbook.sheets.sheet["@attributes"]["sheetId"];
        sheet.file = xlWB.workbook.sheets.sheet["@attributes"]["r:id"].replace("rId", "sheet") + '.xml';
        //1 лист всегда видимый
        sheets.push(sheet);
    }

    wb.Sheets = sheets;
    wb.WBView = xlWB.workbook.bookViews.workbookView["@attributes"];
    if ("activeTab" in xlWB.workbook.bookViews.workbookView["@attributes"]) { } else { wb.WBView.activeTab = 0 };
    //

    return wb;

};

//описание объединенных ячеек
let mcells = [];
function parseMerge(mC) {
    mcells = [];
    let mc = '';
    if (mC) {
        //console.log(mC);
        if (Array.isArray(mC.mergeCell)) {
            mC.mergeCell.forEach(m => {
                mc = m['@attributes']['ref'];
                cells = mc.split(':');

                ic1 = idc(cells[0]);
                ic2 = idc(cells[1]);
                ir1 = idr(cells[0]);
                ir2 = idr(cells[1]);

                for (let i = ir1; i < ir2 + 1; i++) {
                    for (let j = ic1; j < ic2 + 1; j++) {
                        cell = {};
                        //console.log(j + ' ' + i);
                        if (j == ic1 && i == ir1) {
                            if (ic2 > ic1) { cell.colspan = ic2 - ic1 + 1 }
                            if (ir2 > ir1) { cell.rowspan = ir2 - ir1 + 1 }
                            cell.id = cells[0];
                            cell.enabled = true;
                            cell.merge = mc;

                        } else {
                            cell.id = String.fromCharCode(64 + j) + i;
                            cell.enabled = false;

                        }
                        mcells.push(cell);

                        //console.log(cell);
                    }
                }
            })
        }
        else {
            mc = mC.mergeCell['@attributes']['ref'];
            cells = mc.split(':');

            ic1 = idc(cells[0]);
            ic2 = idc(cells[1]);
            ir1 = idr(cells[0]);
            ir2 = idr(cells[1]);

            for (let i = ir1; i < ir2 + 1; i++) {
                for (let j = ic1; j < ic2 + 1; j++) {
                    cell = {};
                    //console.log(j + ' ' + i);
                    if (j == ic1 && i == ir1) {
                        if (ic2 > ic1) { cell.colspan = ic2 - ic1 + 1 }
                        if (ir2 > ir1) { cell.rowspan = ir2 - ir1 + 1 }
                        cell.id = cells[0];
                        cell.enabled = true;
                        cell.merge = mc;

                    } else {
                        cell.id = String.fromCharCode(64 + j) + i;
                        //cell.option = celloption;
                        cell.enabled = false;

                    }
                    mcells.push(cell);
                    //console.log(cell);

                }
            }

        }

    }
    return mcells;
}

let rels = [];
function parseRels(relsjson) {
    rels = [];
    if (Array.isArray(relsjson.Relationships.Relationship)) {
        relsjson.Relationships.Relationship.forEach(rs => {
            if (rs['@attributes']['TargetMode'] == 'External') {
                let rel = {
                    "id": rs['@attributes']['Id'],
                    "target": rs['@attributes']['Target']
                };
                rels.push(rel)
            }
        })
    }
}

//Гиперссылки
function parseHL(hlink) {
    let hcells = [];
    if (hlink && Array.isArray(hlink.hyperlink)) {
        hlink.hyperlink.forEach(hl => {

            rl = rels.find(el => el.id === hl["@attributes"]["r:id"]);
            cell = {
                "id": hl["@attributes"]["ref"],
                "hiperlink": rl.target
            }
            hcells.push(cell);
        })
    }
    return hcells;
}

function toHTML(json, ncols) {

    //объединенные столбцы
    let merges = mcells;
    //Таблица с данными
    let xldata = json;
    let cols = ncols.cols;
    //console.log(cols);

    //let twidth = value.Cols.width;
    let kf = 5.25 * 1.333;//from xml to px - for <td width="" ..>
    let r = 0;
    let kfh = 1.333;//from xml to px - for <th height="" ..>
    let ht = Math.round(15 * kfh);//px высота строки по умолчанию
    let c = 0;
    let wd = Math.round(8.43 * kf);//px ширина колонки по умолчанию

    let widthrow = 0;
    let maxwidth = 0;


    let xlTable = '';

    //колонки листа Excel
    xlTable += `<colgroup><col width=30></col>`;
    icol = 0;
    //console.log(cols);
    cols.forEach(col => {
        xlTable += `<col data-td="td_${icol}" width="${Math.round(col.width * kf)}"></col>`;
        widthrow += Math.round(col.width * kf);
        icol++;
    });

    //не все <col></col> были объявлены в excel
    if (icol < maxcellnumber) {
        //console.log(rownum+' '+(maxcellnumber-c));
        while (icol < maxcellnumber) {
            xlTable += `<col data-td="td_${icol}" width="${wd}"></col>`;
            widthrow += wd;
            icol++;
        }
    }


    xlTable += '</colgroup>';


    //строка с индексами колонок колонок A B C D ...AA AB..
    xlHead = '<thead class="resizable"><tr>';
    xlHead += '<th class="jexcel_selectall"></th>';
    //только двухзначные колонки
    i = 0;
    i1 = 0;
    s1 = '';
    for (var j = 0; j <= maxcellnumber - 1; j++) {
        if (i > 25) {
            //изменить 1 символ
            i = 0; i1++;
            s1 = String.fromCharCode(64 + i1);
        }
        xlHead += `<th data-td="td_${i}">${s1}` + `${String.fromCharCode(65 + i)}<span class="btn btn-right">&nbsp;</span></th>`;//>
        i++;
    };
    xlHead += '</tr></thead>';
    xlTable += xlHead;
    //колонки листа Excel

    xlTable += '<tbody>';
    rownum = 1;
    let delcol = []; //number of deleted columns in [x] row



    xldata.forEach(row => {
        let c = 0;
        widthrow = 0;

        while ((row.r - r) > 1) {
            //Пустой ряд если был пропущен 
            xlTable += `<tr height="${ht}">`;
            //дополнительная колонка с номерами строк
            xlTable += `<td class="row_selectall">${rownum}</td>`;
            //пустые ячейки пустой строки
            for (i = 0; i < maxcellnumber; i++) {
                xlTable += `<td id="${String.fromCharCode(65 + i)}${rownum}" data-td="td_${i}" data-tr="tr_${rownum}"></td>`;
            }
            xlTable += `</tr>`;
            rownum++;
            r++;
        }

        r = row.r;
        if (row.ht) {
            xlTable += `<tr height="${Math.round(row.ht * kfh)}">`;
        } else {
            xlTable += `<tr height="${ht}">`;
        }
        //дополнительная колонка с номерами строк
        xlTable += `<td class="row_selectall">${rownum}</td>`;

        row.columns.forEach(col => {
            ic = idc(col.id);
            while ((ic - c) > 1) {
                //пропущенные ячейки для HTML добавить
                cellid = findMerged(String.fromCharCode(c + 65) + r);
                if (cellid.enabled) {
                    tdattr = '';
                    if (cols[c]) {
                        tdattr += ` width="${Math.round(cols[c].width * kf)}"`;
                        widthrow += Math.round(cols[c].width * kf);
                    } else {
                        //по умолчанию
                        tdattr += ` width="${wd}"`;
                        widthrow += wd;
                    }

                    tdattr += ` id="${String.fromCharCode(c + 65) + r}"`;
                    if (cellid.colspan) { tdattr += ` colspan="${cellid.colspan}"` };
                    if (cellid.rowspan) { tdattr += ` rowspan="${cellid.rowspan}"` };

                    xlTable += `<td ${tdattr} ></td>`;
                }
                c++;
            }

            cellid = findMerged(String.fromCharCode(c + 65) + r);//пока обрабатываются только односимвольные колоки
            //console.log(cols[c]);
            if (cellid.enabled) {
                tdattr = '';
                classattr = '';
                clstyle = '';

                if (col.style) {
                    classattr += `cstyle${col.style}`;
                }
                if (col.type) { classattr += ` ${col.type}` };

                tdattr += ` id="${String.fromCharCode(c + 65) + r}" data-td="td_${c}" data-tr="tr_${r}"`

                if (cellid.colspan) {
                    tdattr += ` colspan="${cellid.colspan}"`
                    //c = c + cellid.colspan - 1;    //увеличить номер занятой колонки
                } else {
                    if (cols[c]) {
                        tdattr += ` width="${Math.round(cols[c].width * kf)}"`;
                        widthrow += Math.round(cols[c].width * kf);
                    }
                    else {
                        //по умолчанию
                        tdattr += ` width="${wd}"`;
                        widthrow += wd;
                    }
                };
                if (cellid.rowspan) { tdattr += ` rowspan="${cellid.rowspan}"` };

                hcl = hiperlinks.find(el => el.id === String.fromCharCode(c + 65) + r);
                //console.log(hcl);

                //значение в ячейке
                if (col.v) {
                    if (hcl) {
                        //hiperlink   
                        if (col.inlineStr) {
                            xlTable += `<td class="${classattr}" ${tdattr} ${clstyle} ><div class="inlineStr"><a href="${hcl.hiperlink}" target="_blank">${col.v}</a></div></td>`;
                        } else {
                            xlTable += `<td class="${classattr}" ${tdattr} ${clstyle} ><a href="${hcl.hiperlink}" target="_blank">${col.v}</a></td>`;
                        }
                    } else {
                        if (col.inlineStr) {
                            xlTable += `<td class="${classattr}" ${tdattr} ${clstyle} ><div class="inlineStr">${col.v}</div></td>`;
                        } else {
                            xlTable += `<td class="${classattr}" ${tdattr} ${clstyle} >${col.v}</td>`;
                        }
                    }
                } else {
                    xlTable += `<td class="${classattr}" ${tdattr} ${clstyle} ></td>`;
                }
            }

            c++;
        })

        //дозаполнить строку ячейками для красивого отображения

        if (c < maxcellnumber) {
            //console.log(rownum + ' ' + (maxcellnumber - c));
            for (i = c; i < (maxcellnumber); i++) {
                xlTable += `<td id="${String.fromCharCode(65 + i)}${rownum}" data-td="td_${i}" data-tr="tr_${rownum}"></td>`;
            }
        }

        xlTable += '</tr>';
        rownum++;
        maxwidth = (maxwidth < widthrow) ? widthrow : maxwidth;
    });

    //Start
    xlTable = `<table id="sheetjs" class="sheetjs" width="${maxwidth}">` + xlTable;
    //END    
    xlTable += '</table>';
    return xlTable;
}

function make_xlsx_lib(MXLSX) {
    MXLSX.version = '0.1.0';

    async function readZIP(file) {

        return new Promise((resolve, reject) => {

            let zip = new JSZip();
            let reader = new FileReaderEx();

            (async () => {
                const buffer1 = await new FileReaderEx().readAsArrayBuffer(file);

                let data = new Uint8Array(buffer1);

                await zip.loadAsync(data);

                //Styles
                let xmlfile = await zip.file("xl/styles.xml").async("string");
                let XmlNode = new DOMParser().parseFromString(xmlfile, 'text/xml');
                let stylesjson = xmlToJson(XmlNode);

                //numFmt for Cells
                cellXfs = parseStyle(stylesjson);
                files.StyleCSS = parseStyleCSS(stylesjson);

                //sharedStrings
                let filename = await zip.file("xl/sharedStrings.xml");
                if (filename) {
                    xmlfile = await zip.file("xl/sharedStrings.xml").async("string");
                    XmlNode = new DOMParser().parseFromString(xmlfile, 'text/xml');
                    //console.log(xmlToJson(XmlNode));
                    sharedStrings = xmlToJson(XmlNode).sst.si;// код...
                    if (!Array.isArray(sharedStrings)) {
                        sharedStrings[0]=xmlToJson(XmlNode).sst.si;
                    }
                } else {
                    sharedStrings = []
                }
                //console.log(sharedStrings);

                //Relationship only Sheet1
                filename = await zip.file("xl/worksheets/_rels/sheet1.xml.rels");
                if (filename) {
                    xmlfile = await zip.file("xl/worksheets/_rels/sheet1.xml.rels").async("string");
                    XmlNode = new DOMParser().parseFromString(xmlfile, 'text/xml');
                    let relsjson = xmlToJson(XmlNode);
                    //console.log(relsjson);
                    parseRels(relsjson);
                }
                // else {
                //     Relationship = []
                // }
                //console.log(rels);


                //console.log(sharedStrings);

                //???
                // xmlfile = await zip.file("xl/workbook.xml").async("string");
                // XmlNode = new DOMParser().parseFromString(xmlfile, 'text/xml');
                // files.Workbook = parseWB(xmlToJson(XmlNode));

                //only Sheet1  
                xmlfile = await zip.file("xl/worksheets/sheet1.xml").async("string");
                XmlNode = new DOMParser().parseFromString(xmlfile, 'text/xml');
                sheet = xmlToJson(XmlNode);
                nsheet = parseSheet(sheet.worksheet.sheetData, sharedStrings, cellXfs);
                files.Sheets = { "sheet1": nsheet };

                //??? так надо, 
                //let maxcell = maxcellnumber;
                //console.log(maxcellnumber);
                // if ('dimension' in sheet.worksheet) {
                //     dimensions = sheet.worksheet.dimension['@attributes']['ref'].split(':');
                //     maxcellnumber = idc(dimensions);
                // }

                if ("cols" in sheet.worksheet) {
                    cols = sheet.worksheet.cols;
                    ncols = parseCols(cols, maxcellnumber);
                    //files.Cols = ncols;
                } else {
                    ncols = { "cols": [] };
                    //files.Cols = ncols;
                }

                mergecells = sheet.worksheet.mergeCells;
                nmerge = parseMerge(mergecells);
                //files.MergeCells = nmerge;

                relscells = sheet.worksheet.hyperlinks;

                hiperlinks = parseHL(relscells);
                // console.log(hiperlinks);

                //files.HTML = toHTML(nsheet, ncols);
                files.HTML = toHTML(nsheet, ncols);

                //console.log(files);
                resolve(files);

            })();
            reader.readAsArrayBuffer(file);
        });

    }

    MXLSX.readZIP = readZIP;

}

make_xlsx_lib(MXLSX);