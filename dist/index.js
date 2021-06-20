"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
var xlsx_1 = __importDefault(require("xlsx"));
var xlsx_style_1 = __importDefault(require("xlsx-style"));
var person1 = {
    name: 'Santiago Gatius',
    age: 28,
    born: new Date(1993, 0, 7),
};
var person2 = {
    name: "Xavier Ruiz",
    age: 28,
    born: new Date(1993, 0, 1)
};
var toArray = function (data) {
    return [data.name, data.age.toString(), data.born.toLocaleDateString('es-ES')];
};
var array = [toArray(person1), toArray(person2)];
var sheet = xlsx_1.default.utils.aoa_to_sheet(array);
console.log(sheet);
var range = xlsx_1.default.utils.decode_range(sheet['!ref']);
console.log(range);
for (var r = 0; r <= range.e.r; r++) {
    for (var c = 0; c <= range.e.c; c++) {
        var cellAddress = { r: r, c: c };
        console.log(cellAddress);
        var cellRef = xlsx_1.default.utils.encode_cell(cellAddress);
        sheet[cellRef].s = {
            fill: {
                fgColor: { rgb: 'FFA3F4B1' } // Add background color
            },
        };
        var cell = sheet[cellRef];
        console.log(cell);
    }
}
var workbook = xlsx_1.default.utils.book_new();
xlsx_1.default.utils.book_append_sheet(workbook, sheet, "Prova");
//XLSX.writeFile(workbook, "output.xlsx", {cellStyles: true});
xlsx_style_1.default.writeFile(workbook, "output.xlsx", { bookType: 'xlsx', bookSST: false, type: 'binary' });
var provaWB = xlsx_1.default.readFile('prova.xlsx', { cellStyles: true });
var sheetName = provaWB.SheetNames[0];
var sheetProva = provaWB.Sheets[sheetName];
var rangeProva = xlsx_1.default.utils.decode_range(sheetProva['!ref']);
for (var row = 0; row <= rangeProva.e.r; row++) {
    for (var col = 0; col <= rangeProva.e.c; col++) {
        var cellAddress = { r: row, c: col };
        var cellRef = xlsx_1.default.utils.encode_cell(cellAddress);
        var cell = sheetProva[cellRef];
        console.log("Celda lectura", cell);
    }
}
