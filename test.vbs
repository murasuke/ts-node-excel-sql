Dim cn, ExcelPath

ExcelPath = "C:\Users\t_nii\Documents\git\activex\ts-node-excel-sql\sample-data.xlsx"
Set cn = CreateObject("ADODB.Connection")
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ExcelPath & ";Extended properties=""Excel 12.0;HDR=YES;"""

Set cmd = CreateObject("ADODB.Command")
cmd.ActiveConnection = cn
cmd.CommandText = "SELECT * FROM [Sheet1$]"
set rs = cmd.Execute()

WScript.Echo rs.Fields.count

Do Until rs.EOF
  WScript.Echo rs("COL1").Value  & ":" & rs("COL2").Value & ":" & rs("COL3").Value
  'レコードセットのカレント行を次に移動
  rs.MoveNext
Loop

  ' レコードセットをClose
rs.Close
cn.Close