<%@ Language=VBScript CODEPAGE=950 %>
<%
	'讀取人員姓名
	p_wk_item_old=trim(request("wk_item_old"))
	p_wk_item_new=trim(request("wk_item_new"))
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
strSQL_show="Select * from " & tb_name & " where wk_item like '"& p_wk_item_old &"' order by doing_date1 desc"
rstObj1.open strSQL_show,conDB,1,3
totalput=rstObj1.recordcount
rstobj1.MoveFirst
if totalput=0 then
else
      '修改資料
      for kj=1 to totalput
            rstObj1.fields("wk_item")= p_wk_item_new                                '主旨
      	'移到下一筆記錄
      		rstObj1.MoveNext
      		if rstObj1.EOF=True then exit for
      next
      rstObj1.UpdateBatch
end if
'關閉資料集
rstObj1.Close
'重設資料變數
set rstObj1=Nothing
'關閉資料庫
conDB.Close
'重設物件變數
set conDB=Nothing 

strURL1="wk_query_oki.asp?q_text="&p_wk_item_new
response.redirect(strURL1)
%>

</body>
</html>
