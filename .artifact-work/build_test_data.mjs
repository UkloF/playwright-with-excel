import fs from 'node:fs/promises';
import {fileURLToPath} from 'node:url';
import {FileBlob, SpreadsheetFile} from '@oai/artifact-tool';

const root = new URL('../', import.meta.url);
const source = new URL('test_data.xlsx', root);
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(fileURLToPath(source)));
const original = await workbook.render({sheetName: 'test_data', range: 'A1:H3', scale: 1, format: 'png'});
await fs.writeFile(new URL('before.png', import.meta.url), new Uint8Array(await original.arrayBuffer()));
try { await fs.access(new URL('test_data.original.xlsx', import.meta.url)); }
catch { await fs.copyFile(source, new URL('test_data.original.xlsx', import.meta.url)); }

const headers = ['testcase name', 'username', 'password', 'items_to_add', 'first_name', 'last_name', 'postal_code', 'expected_result'];
const success = 'Thank you for your order!';
const rows = [
  ['TC01_Purchase_Multiple_Items','standard_user','secret_sauce','Sauce Labs Backpack, Sauce Labs Bike Light, Sauce Labs Bolt T-Shirt','Suda','Meesuk','123',success],
  ['TC02_Empty_First_Name','standard_user','secret_sauce','Sauce Labs Backpack',null,'Jaidee','10110','Error: First Name is required'],
];
const sheet = workbook.worksheets.getItem('test_data');
for (const table of [...sheet.tables.items]) table.delete();
sheet.getRange('A1:H12').clear({applyTo:'all'});
sheet.getRange('A1:H3').values = [headers, ...rows];
sheet.showGridLines = false;
sheet.getRange('A1:H3').format.font = {name:'Arial', size:11};
sheet.getRange('A1:H3').format.verticalAlignment = 'center';
sheet.getRange('A1:H3').format.wrapText = true;
sheet.getRange('A1:H3').setNumberFormat('@');
sheet.getRange('A1:H1').format.fill = '#203A57';
sheet.getRange('A1:H1').format.font = {name:'Arial', size:11, bold:true, color:'#FFFFFF'};
sheet.getRange('A1:H1').format.horizontalAlignment = 'center';
sheet.getRange('A1:H1').format.rowHeight = 30;
sheet.getRange('A2:H3').format.rowHeight = 54;
for (const [col,width] of [['A',34],['B',20],['C',20],['D',48],['E',15],['F',15],['G',16],['H',65]]) {
  sheet.getRange(`${col}1:${col}3`).format.columnWidth = width;
}
for (let r=2;r<=3;r++) {
  sheet.getRange(`A${r}:H${r}`).format.fill = r%2===0 ? '#F0F4F8' : '#FFFFFF';
}
sheet.freezePanes.freezeRows(1);
sheet.tables.add('A1:H3', true, 'ScenarioCases');
workbook.recalculate();
console.log((await workbook.inspect({kind:'table', range:'test_data!A1:H3', tableMaxRows:3, tableMaxCols:8, maxChars:1500})).ndjson);
const preview = await workbook.render({sheetName:'test_data',range:'A1:H3',scale:1,format:'png'});
await fs.writeFile(new URL('after.png', import.meta.url), new Uint8Array(await preview.arrayBuffer()));
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(fileURLToPath(source));
console.log(`Saved ${rows.length} test cases with ${headers.length} columns.`);
