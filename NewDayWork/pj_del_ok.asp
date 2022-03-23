<%@ Language=VBScript CODEPAGE=950 %>
<%
'=====================基本資料=============================		

'專案id
	if request("p00")<>"" then 
		p_00=trim(request("p00"))
	else
		p_00=""
	end if
'專案名稱
	'if request("p02")<>"" then 
		'p_02=trim(request("p02"))
	'else
		'p_02=""
	'end if

'將工作中之專案名稱刪除 
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
	strSQL_show="Select * from " & tb_name & " where pj_id ="& p_00
	rstObj1.open strSQL_show,conDB,1,3
	totalput=rstObj1.recordcount
	if totalput=0 then	
	else
		rstobj1.MoveFirst
		for i=1 to totalput
			'讀取資料
			rstObj1.fields("pj_id")= null
			rstObj1.fields("pj_02")= ""
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

'將project_data中之專案名稱刪除 
	' 連結Access資料庫daywork.mdb
	DBpath=Server.MapPath("./database/daywork.mdb")
	strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
	'建立資料庫連結物件
	set conDB= Server.CreateObject("ADODB.Connection")
	'連結資料庫	
	conDB.Open strCon
	'開啟資料表名稱
	tb_name="project_data"
	'刪除資料之SQL指令字串
	strSQL_del="Delete from " & tb_name & " where pj_id =" & p_00
	'執行SQL指令
	conDB.Execute strSQL_del
	'關閉資料庫 
	conDB.Close
	'重設物件變數 
	set conDB=Nothing 

str_url="pj_list.asp"
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
