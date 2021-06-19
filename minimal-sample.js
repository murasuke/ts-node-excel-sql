//import path from 'path';
const path = require('path');
require('winax');

// Microsoft.ACE.OLEDBプロバイダでExcelを開く
const excelPath = path.join(__dirname, 'sample-data.xlsx');
const connString = `Provider=Microsoft.ACE.OLEDB.12.0;Data Source=${excelPath};Extended Properties="Excel 12.0 Xml;HDR=YES;"`;
const cn = new ActiveXObject('ADODB.Connection');
cn.Open(connString);

const cmd = new ActiveXObject('ADODB.Command');
cmd.ActiveConnection = cn;

// Updateで更新する
cmd.CommandText = `UPDATE [Sheet1$] SET COL2 = 'update_by_node' WHERE COL1 = 3`;
cmd.Execute();

// 更新されていることをSelectで確認
cmd.CommandText = `SELECT * FROM [Sheet1$] WHERE COL1 = 3`;
const rs = cmd.Execute();
console.log(rs.Fields('COL2').Value);

cn.Close();
