/**
 * node.jsでExcelファイルにSQLで更新するサンプル
 * ・`winax`(https://www.npmjs.com/package/winax)でADO(ActiveX Data Object)を利用します
 * ・「Microsoft.ACE.OLEDB」(旧Jetエンジン)プロバイダを利用して、クエリを実行します
 * ・「シート名」が「テーブル名」です。シート名の最後に'$'を追加し、[]で括ります
 * ・1行目が「列名」になります(HDR=YES)
 * ・Select, Insert, Updateが可能です。Deleteはプロバイダがサポートしていません
 */

import fs from 'fs';
import path from 'path';

require('winax');

// 何度も実行できるようにするため、テンプレートをコピーしたファイルに対して更新操作を行う
const xlsxTemplate = './sample-data-template.xlsx';
const xlsxTarget = './sample-data.xlsx';

// ファイルがあれば削除してから、テンプレートをコピー
if (fs.existsSync(xlsxTarget)) { fs.unlinkSync(xlsxTarget); }
fs.copyFileSync(xlsxTemplate, xlsxTarget);

// ADOでExcelに接続
const cn = connectExcel(path.join(__dirname, xlsxTarget));

// 変更前ファイルをSelect(3行)
const resultSet = selectExcel(cn, 'Sheet1');
console.log('変更前データ--------------');
showResultSet(resultSet);

// 3行目を更新、4行目をInsert
updateExcel(cn);
insertExcel(cn);
// deleteExcel(cn); //  OLE DB providerがdeleteをサポートしていないため実行不可

// 変更後データを表示する(4行)
const resultSet2 = selectExcel(cn, 'Sheet1');
console.log('変更後データ--------------');
showResultSet(resultSet2);


cn.Close();

function connectExcel(excelPath: string): ADODB.Connection {
  const connString = `Provider=Microsoft.ACE.OLEDB.12.0;Data Source=${excelPath};Extended Properties="Excel 12.0 Xml;HDR=YES;"`;
  
  const cn = new ActiveXObject('ADODB.Connection');
  cn.Open(connString);
  return cn;
}

function selectExcel(con: ADODB.Connection, tableName: string): [{ [index:string]:string }] {
  const sql = `SELECT * FROM [${tableName}$]`;
  const cmd = createCommand(con, sql);
  const rs = cmd.Execute();
  
  const result: [{[index:string]:string}] = [{}];
  
  for (let rowIndex = 0; !rs.EOF; rowIndex++) {
    let record: {[index:string]:string} = {};
    for( var colIndex = 0; colIndex < rs.Fields.Count; colIndex++ ) {
      const colName = rs.Fields(colIndex).Name;
      const value = rs.Fields(colName).Value;  
      record[colName] = value;
    }
    result.push(record);
    rs.MoveNext();
  }

  return result;
}

function updateExcel(con: ADODB.Connection) {
  const sql = `UPDATE [Sheet1$] SET COL2 = 'update_by_node' WHERE COL1 = 3`;
  const cmd = createCommand(con, sql);
  cmd.Execute();
}

function insertExcel(con: ADODB.Connection) {
  const sql = `INSERT INTO [Sheet1$] VALUES(4, 'insert_by_node', '${(new Date()).toLocaleString("ja")}')`;
  const cmd = createCommand(con, sql);
  cmd.Execute();
}

/**
 * ADODB.Command生成ヘルパ
 * @param con 
 * @param sql 
 * @returns 
 */
function createCommand(con: ADODB.Connection, sql: string): ADODB.Command {
  const cmd = new ActiveXObject('ADODB.Command');
  cmd.ActiveConnection = con;
  cmd.CommandText = sql;
  return cmd;
}

/**
 * ADODB.Recordsetをディクショナリの配列に変換するヘルパ
 * @param resultSet 
 */
function showResultSet(resultSet: [{ [index:string]:string }]) {
  for(let record of resultSet) {
    let str = '';
    for(let key in record) {
      record[key];
      str += `${key}: ${record[key]} `;
    }
    console.log(str);
  }
}