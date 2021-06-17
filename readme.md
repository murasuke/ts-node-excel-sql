# node.js上でSQLをExcelファイルに実行する試み

## 前書き

* Excelのシートをテーブル名、各シートの1行目を列名として`Select`,`Insert`,`Update`,`Delete`をしてみたい

* Windows限定ネタなので

## 実現方法

* ADO(ActiveX Data Object) + Microsoft.ACE.OLEDB.12.0 でExcelファイルを読み書きする


