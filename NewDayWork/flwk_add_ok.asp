<%@ Language=VBScript CODEPAGE=950 %>
<!-- 開啟資料庫 -->
<%
' 連結Access資料庫daywork.mdb
DBpath=Server.MapPath("./database/wk_file.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="wk_file"
'
wk_id=request("wk_id") 
fl_name=request("filename")

'新增資料之SQL指令字串
strSQL_add="Insert into "&tb_name&" ("
strSQL_add=strSQL_add & "wk_id,"
strSQL_add=strSQL_add & "fl_name) values ('"
strSQL_add=strSQL_add & wk_id &"','"
strSQL_add=strSQL_add & fl_name &"')"

'執行SQL指令
conDB.Execute strSQL_add


'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " order by fw_id desc"
rstObj1.open strSQL_show,conDB,3,3
totalput=rstObj1.recordcount
	'列出資料項目
rstobj1.MoveFirst
fw_id1=rstObj1.fields("fw_id")

'關閉資料集
rstObj1.Close
'重設資料變數 
set rstObj1=Nothing
'關閉資料庫 
conDB.Close
'重設物件變數 
set conDB=Nothing 

'str_url="flwk_display.asp?fw_id="&fw_id1

str_url="wk_pj_show.asp?wk_id="&wk_id
response.redirect(str_url) 


%>






<html>
<head>
<title>資料完整新增</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'標楷體';background-color:'##FFFFcc'}
--></style>
</head>
<body>
<center>

</body>
</html>
