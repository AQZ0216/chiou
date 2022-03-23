<%@ Language=VBScript CODEPAGE=950 %>
<%
'修改派工人員
'p_order_old="羽婷"
'p_order_new="Ellie"
p_order_old=request("p_order_old")
p_order_new=request("p_order_new")

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

'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_order like'%" & p_order_old &"%' order by wk_id asc"
rstObj1.open strSQL_show,conDB,3,3
'計算資料總數
totalput=rstObj1.recordcount
if totalput=0 then
else
   rstObj1.MoveFirst
   for j=1 to totalput
      '修改資料
      rstObj1.fields("wk_order")= trim(p_order_new)            '派工者
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

資料修改完成。
所有派工者資料【<%=p_order_old%>】改為【<%=p_order_new%>】。
</body>
</html>
