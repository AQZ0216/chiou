<% @codepage=950%>
<%
	'讀取人員姓名
	worker = Session("worker")
	wk_id=Request("wk_id")
%>
<!-- 開啟資料庫 -->
<!-- #Include file = "./include/opendb_daywork.inc" -->
<%
'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_id =" & wk_id
rstObj1.open strSQL_show,conDB,3,1
'讀取資料
undo_date1=rstObj1.fields("undo_date1")
doing_date1=rstObj1.fields("doing_date1")
done_date1=rstObj1.fields("done_date1")
wk_item=rstObj1.fields("wk_item")
wk_content=rstObj1.fields("wk_content")
wk_order=rstObj1.fields("wk_order")
wk_doer=rstObj1.fields("wk_doer")
wk_checker=rstObj1.fields("wk_checker")
wk_undoer=rstObj1.fields("wk_undoer")
wk_class=rstObj1.fields("wk_class")
if wk_class<>"" then

else
	wk_class="未分類"
end if
wk_group=rstObj1.fields("wk_group")
%>
<%
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
<form name="form1" action="" method="post">
<input type=hidden name="wk_id1" value="<%=wk_id%>">
<input type="hidden" name="worker1" value="<%=worker%>">
<!-- #Include file = "./include/toolbar_show_done.inc" -->
<!-- #Include file = "./include/wk_show_form_done.inc" -->
<hr>
<a href="wk_done_undo.asp?wk_id=<%=wk_id%>" >[<%=worker%>]取消完成</a>
</form>
<center>
</body>
</html>
