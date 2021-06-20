import XLSX, { CellAddress, CellObject, Range, WorkBook, WorkSheet } from 'xlsx';
import XLSXSTYLE from 'xlsx-style';

interface Person {
    name: string;
    age: number;
    born: Date;
}

const person1: Person = {
    name: 'Santiago Gatius',
    age: 28,
    born: new Date(1993, 0, 7),
}

const person2: Person = {
   name: "Xavier Ruiz",
    age: 28,
    born: new Date(1993, 0, 1) 
}

const toArray = (data: Person): string[] => {
    return [data.name, data.age.toString(), data.born.toLocaleDateString('es-ES')];

}

const array: string[][] = [toArray(person1), toArray(person2)];


const sheet: WorkSheet = XLSX.utils.aoa_to_sheet(array);
console.log(sheet)
const range: Range = XLSX.utils.decode_range(sheet['!ref'] as string);
console.log(range);
for( let r = 0; r <= range.e.r; r++){
    for(let c = 0; c <= range.e.c; c++){
        const cellAddress: CellAddress = {r, c};
        console.log(cellAddress);
        const cellRef: string = XLSX.utils.encode_cell(cellAddress);
        (sheet[cellRef] as CellObject).s = {
            fill: {
                fgColor: { rgb: 'FFA3F4B1' } // Add background color
            },
        };
        const cell: CellObject = sheet[cellRef];
        
        console.log(cell);
    }
}
const workbook: WorkBook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, sheet, "Prova");
//XLSX.writeFile(workbook, "output.xlsx", {cellStyles: true});
XLSXSTYLE.writeFile(workbook, "output.xlsx", {bookType: 'xlsx', bookSST: false, type: 'binary'});

const provaWB: WorkBook = XLSX.readFile('prova.xlsx', {cellStyles: true});
const sheetName: string = provaWB.SheetNames[0];
const sheetProva: WorkSheet = provaWB.Sheets[sheetName];
const rangeProva: Range = XLSX.utils.decode_range(sheetProva['!ref'] as string);
for(let row = 0; row <= rangeProva.e.r; row++){
    for(let col = 0; col <= rangeProva.e.c; col++){
        const cellAddress: CellAddress = {r: row, c: col};
        const cellRef: string = XLSX.utils.encode_cell(cellAddress);
        const cell: CellObject = sheetProva[cellRef];
        console.log("Celda lectura",cell);
    }
}