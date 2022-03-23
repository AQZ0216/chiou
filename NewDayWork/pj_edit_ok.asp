<%@ Language=VBScript CODEPAGE=950 %>
<%
'=====================基本資料=============================		
'專案id
	if request("p_id")<>"" then 
		pj_id=request("p_id")
	else
		pj_id=""
	end if

'專案編號
	if request("p01")<>"" then 
		p_01=trim(request("p01"))
	else
		p_01=""
	end if
'專案名稱
	if request("p02")<>"" then 
		p_02=trim(request("p02"))
	else
		p_02=""
	end if
'專案說明p_03
	if request("p03")<>"" then 
		p_03=trim(request("p03"))
	else
		p_03=""
	end if

%> 
<!-- 開啟資料庫 -->
<%
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
	strSQL_show="Select * from " & tb_name & " where pj_id="& pj_id 
	rstObj1.open strSQL_show,conDB,1,3
	rstobj1.MoveFirst
	'讀取資料
	rstObj1.fields("pj_01")= trim(p_01)
	rstObj1.fields("pj_02")= trim(p_02)
	rstObj1.fields("pj_03")= trim(p_03)
	rstObj1.UpdateBatch
	'關閉資料集
	rstObj1.Close
	'重設資料變數 
	set rstObj1=Nothing

'關閉資料庫 
conDB.Close
'重設物件變數 
set conDB=Nothing 

	'將工作中pj_02專案名稱更正 
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
	strSQL_show="Select * from " & tb_name & " where pj_id="& pj_id 
	rstObj1.open strSQL_show,conDB,1,3
	'計算資料總數	
	totalput=rstObj1.recordcount
	if totalput=0 then
	else
		'列出資料項目
		rstobj1.MoveFirst
		for j=1 to totalput
			rstObj1.fields("pj_02")= trim(p_02)'專案名稱更正
			'移到下一筆記錄
			rstObj1.MoveNext
			if rstObj1.EOF=True then exit for
		next
	end if
	rstObj1.UpdateBatch
	'關閉資料集
	rstObj1.Close
	'重設資料變數 
	set rstObj1=Nothing
	'關閉資料庫 
	conDB.Close
	'重設物件變數 
	set conDB=Nothing 

str_url="pj_show.asp?p_id="&pj_id
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
