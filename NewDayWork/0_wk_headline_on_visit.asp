<% @codepage=950%>
<%
	'讀取人員姓名
	worker = Session("worker")
%>

<%
      '將工作列為重大訊息
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
      strSQL_show="Select * from " & tb_name & " where wk_item like '%來訪%' or wk_item like '%拜訪%' or wk_item like '%到公司%'"
      rstObj1.open strSQL_show,conDB,1,3
	'計算資料總數	
	totalput=rstObj1.recordcount
	if totalput= 0 then
	else
		'移至第一筆資料
		rstObj1.MoveFirst
	    for kj=1 to totalput      
      	rstObj1.fields("headline")=10
      	rstObj1.UpdateBatch
	      '移到下一筆記錄
	      rstObj1.MoveNext
	      if rstObj1.EOF=True then exit for
	    next
	end if
      '關閉資料集
      rstObj1.Close
      '重設資料變數
      set rstObj1=Nothing
      '關閉資料庫
      conDB.Close
      '重設物件變數
      set conDB=Nothing
%>

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
	已將拜訪資料列為重大訊息。
</center>
</body>
</html>
