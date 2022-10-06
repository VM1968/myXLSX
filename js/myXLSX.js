
let MXLSX = {};

class FileReaderEx extends FileReader {
    constructor() {
        super();
    }

    #readAs(blob, ctx) {
        return new Promise((res, rej) => {
            super.addEventListener("load", ({ target }) => res(target.result));
            super.addEventListener("error", ({ target }) => rej(target.error));
            super[ctx](blob);
        });
    }
    #readAsEnc(blob, encoding, ctx) {
        return new Promise((res, rej) => {
            super.addEventListener("load", ({ target }) => res(target.result));
            super.addEventListener("error", ({ target }) => rej(target.error));
            super[ctx](blob, encoding);
        });
    }

    readAsArrayBuffer(blob) {
        return this.#readAs(blob, "readAsArrayBuffer");
    }

    readAsDataURL(blob) {
        return this.#readAs(blob, "readAsDataURL");
    }

    readAsText(blob) {
        return this.#readAs(blob, "readAsText");
    }

    readAsTextEnc(blob, encoding) {
        console.log(encoding);
        return this.#readAsEnc(blob, encoding, "readAsText");
    }
}

const files = {};

const position = ["left", "right", "top", "bottom"];

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


function is_datefmt(fmt) {
    const regex = /[msyh]/g;
    if ((m = regex.exec(fmt)) !== null) { return true }
    else { return false }
}

function findMerged(idcell) {
    //находится ли в списке объединяемых ячеек с дополнительными параметрами "rowspan","colspan"
    //иначе обычная видимая ячейка

    let fcell = {};

    fcell=mcells.find(el=>el.id===idcell);

    if (!fcell) {
        fcell={"id" : idcell,"enabled" : true}
    }
    return fcell;
}

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

function idr(str) {
    //id row for cell "A5" = 5
    const regex = /[0-9]{1,}/;
    let m;
    let id = 0;
    m = regex.exec(str);
    id = Number(m[0]);
    return id
}

function parseExcelDate(excelSerialDate) {
    //(Excel_time-DATE(1970,1,1))*86400 с   //*1000 мс
    return new Date(Math.round((excelSerialDate - 25569) * 86400000));
};


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

let charset = [];
charset[0] = "Ansi"
charset[1] = "Default"
charset[2] = "Symbol"
charset[77] = "Mac"
charset[128] = "ShiftJIS"
charset[129] = "Hangeul"
charset[130] = "Johab"
charset[134] = "GB2312"
charset[136] = "ChineseBig5"
charset[161] = "Greek"
charset[162] = "Turkish"
charset[163] = "Vietnamese"
charset[177] = "Hebrew"
charset[178] = "Arabic"
charset[186] = "Baltic"
charset[204] = "Russian"
charset[222] = "Thai"
charset[238] = "EastEurope"
charset[255] = "Oem"

//indexed color c# ARGB Value
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


function parseStyle(xlStyle) {
    //No CSS rule
    let styles = {};

    //NumberFmt
    //Числовые форматы
    let numFmt = [];
    //стандартные числовые форматы
    init_table(numFmt);

    //дополнительные форматы для чисел, заданные в файлк EXCEL
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
    xlStyle.styleSheet.cellXfs.xf.forEach(cxf => {
        cellxf = {};
        cellxf.numFmt = numFmt[cxf["@attributes"]["numFmtId"]];
        cellxfs.push(cellxf);
    })
    styles.CellXf = cellxfs;
    return styles;
}

function parseStyleCSS(xlStyle) {
    //CSS rule
    //Fonts CSS
    let fontsCSS = [];

    function readFont(fnt) {
        fontcss = '';
        fontcss += 'font-family: ' + fnt.name["@attributes"]["val"] + ';';
        fontcss += 'font-size: ' + Number(fnt.sz["@attributes"]["val"]) + 'pt;';

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
    xlStyle.styleSheet.fills.fill.forEach(fl => {
        fillcss = '';

        if ("patternFill" in fl) {
            if ("fgColor" in fl.patternFill) {
                if ("rgb" in fl.patternFill.fgColor["@attributes"]) {
                    fillcss += 'background-color: #' + fl.patternFill.fgColor["@attributes"]["rgb"].slice(2, 8) + ';';

                }
                if ("indexed" in fl.patternFill.fgColor["@attributes"]) {
                    fillcss += 'background-color: #' + indexed[fl.patternFill.fgColor["@attributes"]["indexed"]].slice(2, 8) + ';';
                }
                if ("theme" in fl.patternFill.fgColor["@attributes"]) {
                    fillcss += 'background-color: #' + indexedToRGB(themecolor[fl.patternFill.fgColor["@attributes"]["theme"]], fl.patternFill.fgColor["@attributes"]["tint"]) + ';';
                }
            }
        };


        fillsCSS.push(fillcss);
    })
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
                        bordercss += 'border-' + ps + ': solid ' + brd[ps]["@attributes"]["style"] + clr + ';'
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
    classCSS='<style>';

    //classCSS.push('table{border-collapse: collapse;table-layout: fixed;}');
    // table-layout: fixed;       border: 1px solid #e6e6e6;
    //htstyle.push('td{border: 1px solid #e6e6e6;}');    
    classCSS+='table{border-collapse: collapse;table-layout: fixed;}';

    //classCSS.push('td{padding-top:1px;padding-right:1px;padding-left:1px;vertical-align: bottom;white-space: nowrap;}');
    classCSS+='td{padding-top:1px;padding-right:1px;padding-left:1px;vertical-align: bottom;white-space: nowrap;}'

    //classCSS.push('td.Number {text-align: end;}');
    classCSS+='td.Number {text-align: end;}';
    //    console.log(xlStyle);

    xlStyle.styleSheet.cellXfs.xf.forEach(function (item, ind) {
        console
        strstyle = '.cstyle' + ind + ' {';
        strstyle += fontsCSS[item["@attributes"]["fontId"]];
        strstyle += fillsCSS[item["@attributes"]["fillId"]];
        strstyle += bordersCSS[item["@attributes"]["borderId"]];

        if ('alignment' in item && 'horizontal' in item.alignment["@attributes"]) { strstyle += ' text-align:' + item.alignment["@attributes"].horizontal + ' !important;' }
        if ('alignment' in item && 'vertical' in item.alignment["@attributes"]) {
            if (item.alignment["@attributes"].vertical == 'center') { strstyle += ' vertical-align: middle;' }
            else { strstyle += ' vertical-align:' + item.alignment["@attributes"].vertical + ';' }
        }
        if ('alignment' in item && 'wrapText' in item.alignment["@attributes"]) { strstyle += ' white-space: normal;' }

        strstyle += '}';
        //classCSS.push(strstyle);
        classCSS+=strstyle;
    })

    //console.log(classCSS);
    classCSS+='</style>';
    return classCSS;
}

function parseSheet(XLsheet, sst, styles) {
    let xlsheet = XLsheet;
    let rows = [];
    let mcolumn = {};

    function readCell(column) {
        switch (column["@attributes"]["t"]) {
            case "s":
                //строка общая в отдельном массиве
                //console.log(sst[column.v["#text"]]);
                mcolumn = {
                    "id": column["@attributes"]["r"],
                    "v": sst[column.v["#text"]]["t"]["#text"]
                };
                if ("s" in column["@attributes"]) {
                    mcolumn.style = column["@attributes"]["s"]//styles.CellXf[column["@attributes"]["s"]]
                }
                if ("@attributes" in sst[column.v["#text"]]["t"] && "xml:space" in sst[column.v["#text"]]["t"]["@attributes"]) {
                    mcolumn.whitespace = 'pre';
                }
                break;
            case "inlineStr":
                //строка непосредственно в ячейке
                mcolumn = {
                    "id": column["@attributes"]["r"],
                    "v": column.is.t["#text"]
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
                    "type" : "Number"
                };
                if ("s" in column["@attributes"]) {
                    mcolumn.style = column["@attributes"]["s"]//styles.CellXf[column["@attributes"]["s"]]
                }
                break;
        }
        c.push(mcolumn);
    }

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
        row.columns = c;
        rows.push(row);
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

function parseCols(Cols, maxcell) {
    //console.log(Cols);
    let xlcols = {};
    let cols = [];
    //let width = 0;

    if (Array.isArray(Cols.col)) {
        Cols.col.forEach(cl => {
            let i = 0;
            //console.log(cl);
            for (i = Number(cl["@attributes"]["min"]); i <= Number(cl["@attributes"]["max"]); i++) {
                //console.log(i);
                col = {};
                col.width = cl["@attributes"]["width"];
                //    width = col.width;
                cols[i - 1] = col;
                if (i > maxcell) { break };
            }
        })
    }
    else {
        //авто ширина колонок?
        for (let i = Number(Cols.col["@attributes"]["min"]); i <= Number(Cols.col["@attributes"]["max"]); i++) {
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

let mcells = [];
function parseMerge(mC) {
    mcells=[];
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

function toHTMLstr(json, ncols) {

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

    //let xltable = document.createElement('table');
    // xltable.setAttribute('border',  '0');
    // xltable.setAttribute('cellpadding',  '0');
    // xltable.setAttribute('cellspacing',  '0');

    //let xlTable = [];
    let xlTable = '';
    // xlTable.push('<table>');

    //let htcolgroup = document.createElement('colgroup');
    //xlTable.push('<colgroup>');
    xlTable+='<colgroup>';
    
    cols.forEach(col => {
        //xlTable.push(`<col width="${Math.round(col.width * kf)}"></col>`);
        xlTable+=`<col width="${Math.round(col.width * kf)}"></col>`;
    });
    //xlTable.push('</colgroup>');
    xlTable+='</colgroup>';

    //xlTable.push('<tbody>');
    xlTable+='<tbody>';

    xldata.forEach(row => {
        c = 0;
        widthrow = 0;
        while ((row.r - r) > 1) {
            //Add default row
            //xlTable.push(`<tr height="${ht}"></tr>`);
            xlTable+=`<tr height="${ht}"></tr>`;
            r++;
        }

        r = row.r;
        if (row.ht) {
            //xlTable.push(`<tr height="${Math.round(row.ht * kfh)}">`);
            xlTable+=`<tr height="${Math.round(row.ht * kfh)}">`;
        } else {
            //xlTable.push(`<tr height="${ht}">`);
            xlTable+=`<tr height="${ht}">`;
        }
        
        row.columns.forEach(col => {

            ic = idc(col.id);
            while ((ic - c) > 1) {
                //пропущенные ячейки для HTML добавить
                cellid = findMerged(String.fromCharCode(c + 65) + r);   
                if (cellid.enabled) {
                    tdattr='';
                    if (cols[c]) {
                        tdattr += ` width="${Math.round(cols[c].width * kf)}"`;
                        widthrow += Math.round(cols[c].width * kf);
                    } else {
                        //по умолчанию
                        tdattr += ` width="${wd}"`;
                        widthrow += wd;
                    }
                   
                   tdattr +=` id="${String.fromCharCode(c + 65) + r}"`;
                    if (cellid.colspan) { tdattr +=` colspan="${cellid.colspan}"`};
                    if (cellid.rowspan) { tdattr +=` rowspan="${cellid.rowspan}"`};
                  
                    //xlTable.push(`<td ${tdattr} ></td>`);
                    xlTable+=`<td ${tdattr} ></td>`;
                }
                c++;
            }

            cellid = findMerged(String.fromCharCode(c + 65) + r);
            //console.log(cols[c]);
            if (cellid.enabled) {
                tdattr ='';
                classattr = '';

                // if ("whitespace" in col && cols[c]) {
                //     clstyle += `max-width: ${Math.round(cols[c].width * kf)}px; white-space: pre;`
                // }

                if (col.style) {
                    //htcell.classList.add("cstyle" + col.style);
                    classattr +=`cstyle${col.style}`;
                }
                if (col.type) {classattr +=` ${col.type}`};

                tdattr +=` id="${String.fromCharCode(c + 65) + r}"`
                
                if (cellid.colspan) {
                    tdattr +=` colspan="${cellid.colspan}"`
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
                if (cellid.rowspan) {  tdattr +=` rowspan="${cellid.rowspan}"`};

                //xlTable.push(`<td class="${classattr}" ${tdattr} >${col.v}</td>`);
                xlTable+=`<td class="${classattr}" ${tdattr} >${col.v}</td>`;
            }
            c++;
        })
        //xlTable.push('</tr>');
        xlTable+='</tr>';
        maxwidth = (maxwidth < widthrow) ? widthrow : maxwidth;
    });

    // xlTable[0]=(`<table width="${maxwidth}">`);
    // xlTable.push('</table>');
//Start
    xlTable=`<table width="${maxwidth}">`+xlTable;
//END    
    xlTable+='</table>';
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
                
                //numFmt for Cells
                cellXfs = parseStyle(xmlToJson(XmlNode));
                files.StyleCSS = parseStyleCSS(xmlToJson(XmlNode));

                //sharedStrings
                filename = await zip.file("xl/sharedStrings.xml");
                if (filename) {
                    xmlfile = await zip.file("xl/sharedStrings.xml").async("string");
                    XmlNode = new DOMParser().parseFromString(xmlfile, 'text/xml');
                    sharedStrings = xmlToJson(XmlNode).sst.si;// код...
                } else {
                    sharedStrings = []
                }

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
                let maxcell = 30;

                if ('dimension' in sheet.worksheet) {
                    dimensions = sheet.worksheet.dimension['@attributes']['ref'].split(':');
                    maxcell = idc(dimensions);
                }

                if ("cols" in sheet.worksheet) {
                    cols = sheet.worksheet.cols;
                    ncols = parseCols(cols, maxcell);
                    //files.Cols = ncols;
                } else {
                    ncols = { "cols": [] };
                    //files.Cols = ncols;
                }

                mergecells = sheet.worksheet.mergeCells;
                nmerge = parseMerge(mergecells);
                //files.MergeCells = nmerge;

                //files.HTML = toHTML(nsheet, ncols);
                files.HTMLstr=toHTMLstr(nsheet, ncols);

                //console.log(files);
                resolve(files);

            })();
            reader.readAsArrayBuffer(file);
        });

    }

    MXLSX.readZIP = readZIP;

}

make_xlsx_lib(MXLSX);