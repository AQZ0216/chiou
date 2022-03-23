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
wk_group=rstObj1.fields("wk_group")
wk_exe=rstObj1.fields("wk_exe")
wk_pjn=rstObj1.fields("pj_02")   '專案名稱
%>
<%
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
body{font-family:'微軟正黑體';background-color:'#F0FFF0'}
input{font-family:'微軟正黑體';}
textarea{font-family:'微軟正黑體';}
SELECT{font-family:'微軟正黑體';font-size:12pt;}
td{font-family:'微軟正黑體';}
--></style>
</HEAD>
<BODY>
<center>
<form name="form1" action="" method="post">
<input type="hidden" name="wk_id1" value="<%=wk_id%>">
<input type="hidden" name="worker1" value="<%=worker%>">
<%if wk_group="一般工作" then%>
<!-- #Include file = "./include/toolbar_show.inc" -->
<%else%>
<!-- #Include file = "./include/toolbar_pj_show.inc" -->
<%end if%>
<!-- #Include file = "./include/wk_show_form.inc" -->

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
strSQL_show="Select * from " & tb_name & " where wk_id =" & wk_id &" and del_ok = false order by fl_date desc"
rstObj1.open strSQL_show,conDB,3,1
totalput=rstObj1.recordcount
if totalput=0 then
else
%>
<table border=1 cellspacing=0 cellpadding=0 width=750 bgcolor="#CCEEFF">
<col width=40 style="text-align:center;">
<col width=280 style="padding-left:5px;text-align:left;">
<col width=210 style="padding-left:5px;text-align:left;">
<col width=90 style="text-align:center;">
<tr>
<td colspan=4>附件列表</td>
</tr>
<tr>
<td >序號</td>
<td align=center >檔案說明</td>
<td align=center >檔案名稱  [上傳者]</td>
<td >建檔日期</td>
</tr>
<%
	'列出資料項目
	rstobj1.MoveFirst
	for fi=1 to totalput
	'讀取資料
		pfl_id=rstObj1.fields("fl_id")
		pwk_id=rstObj1.fields("wk_id")
		pfl_name=rstObj1.fields("fl_name")
		pfl_item=rstObj1.fields("fl_item")
		pfl_inputer=rstObj1.fields("fl_inputer")
		pfl_history= rstObj1.fields("fl_history")
		pfl_date=rstObj1.fields("fl_date")
		str_none=pwk_id&"_"
		str_pfl_name=right(pfl_name,len(pfl_name)-len(pwk_id)-1)
%>
<tr>
<td ><%=fi%></td>
<td >
<a href="./1_ulf_item_edit.asp?fl_id=<%=pfl_id%>" target="_self" title="修改檔案說明。" ><img src="./img/change.png" style="vertical-align:middle;height:16px;cursor:hand;border:0;" ></a>
<%=pfl_item%>
</td>
<td ><a href="./file_att/<%=pfl_name%>" target="_blank" title="<%=pfl_history%>"><%=str_pfl_name%></a>  [<%=pfl_inputer%>]</td>
<td ><%=pfl_date%></td>
</tr>
<%
	'移到下一筆記錄
		rstObj1.MoveNext
		if rstObj1.EOF=True then exit for
	next	
%>

</table>
<%
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
</form>
<center>
</body>
</html>
