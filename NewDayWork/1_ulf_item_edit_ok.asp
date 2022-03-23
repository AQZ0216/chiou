<%@ Language=VBScript CODEPAGE=950 %>

<%
	'讀取人員姓名
	worker = Session("worker")
	fl_id=Request("fl_id")
	pfl_item=Request("item")
	p_fl_history=now()&"〔"&worker&"〕修改檔案說明。"
%>
<%
'附加檔案列表
' 連結Access資料庫daywork.mdb
DBpath=Server.MapPath("./database/attach_file.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="file_data"
'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where fl_id =" & fl_id &" and del_ok = false"
rstObj1.open strSQL_show,conDB,3,3
totalput=rstObj1.recordcount
if totalput=0 then
else
	'列出資料項目
	rstobj1.MoveFirst
		pfl_id=rstObj1.fields("fl_id")
		pwk_id=rstObj1.fields("wk_id")
		pfl_name=rstObj1.fields("fl_name")      '檔案名稱
		pfl_date=rstObj1.fields("fl_date")           '建檔日期
		rstObj1.fields("fl_item")=pfl_item           '檔案說明
		rstObj1.fields("fl_history")=rstObj1.fields("fl_history") & chr(13) & p_fl_history '修改過程
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

'response.write "檔案刪除完成"
myURL="wk_show.asp?wk_id="&pwk_id
Response.Redirect (myURL)
%>

<HTML>
<HEAD>
<Title>上傳檔案功能程式</Title>
<META http-equiv="Content-Type" >
<META name="Generator" >
</HEAD>
<BODY>
<center>
檔案刪除完成!!
<hr>
<a href="wk_show.asp?wk_id=<%=pwk_id%>" target="_self">回工作頁面</a>
</center>
</BODY>
</HTML>