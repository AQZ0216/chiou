<%@ Language=VBScript CODEPAGE=950 %>
<%
'=====================基本資料=============================		

'工作編號p_00
	if request("p00")<>"" then 
		p_00=trim(request("p00"))
	else
		p_00=""
	end if
'專案id p_id
	if request("p_id")<>"" then 
		p_id=trim(request("p_id"))
	else
		p_id=null
	end if
'專案名稱
if isnull(p_id) then
	p_02=""
else
	' 連結Access資料庫./database/daywork.mdb
	DBpath=Server.MapPath("./database/daywork.mdb")
	strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
	'建立資料庫連結物件
	set conDB= Server.CreateObject("ADODB.Connection")
	'連結資料庫	
	conDB.Open strCon
	'開啟資料表名稱
	tb_name="project_data"
	'建立資料庫存取物件	
	set rstObj1=Server.CreateObject("ADODB.Recordset")
	strSQL_show="Select * from " & tb_name & " where pj_id="& p_id
	rstObj1.open strSQL_show,conDB,3,3
	p_02=rstObj1.fields("pj_02")	'專案名稱
	'關閉資料集
	rstObj1.Close
	'重設資料變數 
	set rstObj1=Nothing
	'關閉資料庫 
	conDB.Close
	'重設物件變數 
	set conDB=Nothing 
end if

if p_00="" then
else
	'將名稱加入工作中 
	' 連結Access資料庫daywork.mdb
	DBpath=Server.MapPath("./database/daywork.mdb")
	strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
	'建立資料庫連結物件
	set conDB= Server.CreateObject("ADODB.Connection")
	'連結資料庫	
	conDB.Open strCon
	'開啟資料表名稱
	tb_name="work_data"
	'建立資料庫存取物件	
	set rstObj1=Server.CreateObject("ADODB.Recordset")
	strSQL_show="Select * from " & tb_name & " where wk_id="& p_00 
	rstObj1.open strSQL_show,conDB,1,3
	rstobj1.MoveFirst
	'讀取資料
	rstObj1.fields("pj_id")= p_id
	rstObj1.fields("pj_02")= trim(p_02)
	rstObj1.UpdateBatch
	'關閉資料集
	rstObj1.Close
	'重設資料變數 
	set rstObj1=Nothing
	'關閉資料庫 
	conDB.Close
	'重設物件變數 
	set conDB=Nothing 
end if
str_url="pj_show.asp?p_id="&p_id
response.redirect(str_url) 

%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>新增</title>
<style type="text/css"><!--
body{font-family:'新細明體';background-color:'##FFFFcc'}
--></style>
</head>
<body>
<center>

</body>
</html>
