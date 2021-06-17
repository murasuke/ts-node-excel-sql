import fs from 'fs';
import path from 'path';

require('winax');

const filename = 'persons.mdb';
const data_path = path.join(__dirname,  filename);

const xlsxTemplate = './sample-data-template.xlsx';
const xlsxTarget = './sample-data.xlsx';

// ファイルがあれば削除
if (fs.existsSync(xlsxTarget)) {
  fs.unlinkSync(xlsxTarget);
}

fs.copyFileSync(xlsxTemplate, xlsxTarget)

console.log(path.join(__dirname, xlsxTarget));

const connString = `Provider=Microsoft.ACE.OLEDB.12.0;Data Source={${path.join(__dirname, xlsxTarget)}};Extended Properties="Excel 12.0 Xml;HDR=YES;"`;

const cn = new ActiveXObject('ADODB.Connection');
cn.Open(connString)
const cmd = new ActiveXObject('ADODB.Command');
cmd.ActiveConnection = cn
cmd.CommandText = 'SELECT * FROM [Sheet1$]';
const rs = cmd.Execute();

console.log(`Result field count: ${rs.Fields.Count}`);

while (!rs.EOF) {
  const id = rs.Fields('COL1').Value;  
  const name = rs.Fields('COL2').Value;
  const date = rs.Fields('COL3').Value;

  console.log(`> id: ${id} name: ${name} date: ${date} `);
  rs.MoveNext();
}