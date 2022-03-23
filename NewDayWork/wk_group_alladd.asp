<% @codepage=950%>
<%
	'讀取人員姓名
	worker = Session("worker")
	'讀取今天日期
	ckdate=date()+2
%>
<!-- 開啟資料庫 -->
<!-- #Include file = "./include/opendb_daywork.inc" -->

<HTML>
<HEAD>
<title>工作管理系統</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'標楷體';background-color:'#F0FFF0'}
--></style>
</HEAD>
<BODY>
<center>

<%
'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name
rstObj1.open strSQL_show,conDB,3,3
totalput=rstObj1.recordcount
if totalput=0 then

else
	'列出資料項目
	rstobj1.MoveFirst
	for i=1 to totalput
	'讀取資料
		rstObj1.fields("wk_group")="一般工作"
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
%>
<center>
</body>
</html>
