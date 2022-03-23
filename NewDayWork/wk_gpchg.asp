<%@ Language=VBScript CODEPAGE=950 %>
<%
	'讀取人員姓名
	worker = Session("worker")
	wk_id=Request("wk_id")
'	p_wk_content=trim(request("wk_content"))
'	p_wk_item=trim(request("wk_item"))
'	p_doing_date1=request("doing_date1")
	p_wk_group="專案工作"
%>
<html>
<head>
<title>資料修改</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'標楷體';background-color:'##FFFFcc'}
--></style>
</head>
<body>
<center>
<!-- 開啟資料庫 -->
<%
' 連結Access資料庫daywork.mdb
DBpath=Server.MapPath("./database/daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="work_data"
	'----------2003/03/15修正 
	'修改資料之SQL指令字串 全部資料
	'strSQL_edit="Update "&tb_name&" set wk_content='"&request("wk_content")&"'"
	'strSQL_edit=strSQL_edit & ",doing_date1=#"& request("doing_date1") &"#"
	'strSQL_edit=strSQL_edit & ",wk_item='"& request("wk_item") &"'"
	'strSQL_edit=strSQL_edit & " where wk_id =" & wk_id
	'執行SQL指令
	'conDB.Execute strSQL_edit
	'---------------------------------------------------------
'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_id="&wk_id
rstObj1.open strSQL_show,conDB,1,3
rstobj1.MoveFirst
'讀取資料
rstObj1.fields("wk_content")= date()&"轉為專案工作："&chr(13)&rstObj1.fields("wk_content")
'rstObj1.fields("doing_date1")= p_doing_date1
'rstObj1.fields("wk_item")= p_wk_item
rstObj1.fields("wk_group")= p_wk_group
rstObj1.UpdateBatch
'關閉資料集
rstObj1.Close
'重設資料變數 
set rstObj1=Nothing
'關閉資料庫 
conDB.Close
'重設物件變數 
set conDB=Nothing 

strURL1="wk_show.asp?wk_id="&wk_id
response.redirect(strURL1)
%>

</body>
</html>
