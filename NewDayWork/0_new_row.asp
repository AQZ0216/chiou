<%@ Language=VBScript CODEPAGE=950 %>
<%
'轉換日期格式函數
Function FormatSQLDate(dDate)
	Select Case Vartype(dDate)
	Case 7: '日期時間
		FormatSQLDate="#"&FormatDateTime(dDate,2)&"#"
	Case 8: '字串
		If IsDate(dDate) then
			FormatSQLDate="#"&FormatDateTime(dDate,2)&"#"
		Else
			FormatSQLDate="NULL"
		End if
	Case Else
		FormatSQLDate="NULL"
	End select
End function

%>
<!-- 開啟資料庫 -->
<%

'設定讀取資料編號
new_row=Request("new_row")    'p_id

' 連結Access資料庫./database/linkweb.mdb
DBpath=Server.MapPath("./database/linkweb.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="linkdata"
'======================================================================
for zd=1 to 6

'新增資料之SQL指令字串
strSQL_add="Insert into "& tb_name &" (lk_row,lk_col,lk_show) values ('"
strSQL_add=strSQL_add & new_row &"','"
strSQL_add=strSQL_add & zd &"',"

strSQL_add=strSQL_add & true &")"
'執行SQL指令
conDB.Execute strSQL_add
next
'======================================================================

'關閉資料庫 
conDB.Close
'重設物件變數 
set conDB=Nothing
 
nexturl="firstpage_elink.asp"
response.redirect(nexturl)
%>

	
<html>
<head>
<title>確定新增資料</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'新細明體';background-color :'#FFFEEE'}
input{font-family:'新細明體';font-size:12pt;cursor:hand;}
select{font-family:'新細明體';font-size:10pt;cursor:hand;}
--></style>
</head>
<body>
</body>
</html>
