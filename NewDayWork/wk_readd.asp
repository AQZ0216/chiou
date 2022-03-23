<%@ Language=VBScript CODEPAGE=950 %>
<!-- #Include file = "./include/array_worker_crew.inc" -->
<!-- Include file = "./include/array_worker.inc" -->
<!-- #Include file = "./include/workinput.inc" -->
<!-- #Include file = "./misc_data/array_place.inc" -->	
<!-- #Include file = "./misc_data/array_thing.inc" -->	
<!-- #Include file = "./misc_data/array_writer.inc" -->	

<%
	'讀取人員姓名
	worker = Session("worker")
	wk_id=Request("wk_id")
'工作等級陣列 
dim wk_class_a
wk_class_a=array("","A","B","C","D")
wk_class_no=ubound(wk_class_a)+1
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
wk_class1=rstObj1.fields("wk_class")
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
body{font-family:'標楷體';background-color:'#F0FFF0'}
--></style>
</HEAD>
<BODY>
<center>
<form name="form1" action="wk_readd_ok.asp" method="post">
<input type=hidden name='wk_group' value="一般工作" >
<!-- 重新公告表格 -->
<script language=vbscript>
<%
for i=1 TO worker_no
%>
sub worker_s<%=i%>_click
	if document.form1.all_worker.value="" then
		document.form1.all_worker.value=Trim(document.form1.worker_s<%=i%>.value)
	else
	  if instr(1,document.form1.all_worker.value,document.form1.worker_s<%=i%>.value,1)=0 then
		document.form1.all_worker.value=document.form1.all_worker.value &","& Trim(document.form1.worker_s<%=i%>.value)
	  end if
	end if
end sub
<%
next
%>
sub all_sel_click
	document.form1.all_worker.value=""
	<%
	for i=1 TO worker_no
	%>	
		worker_s<%=i%>_click
	<%
	next
	%>	
end sub
sub all_unsel_click
	'document.form1.all_worker.value=document.form1.worker1.value
	document.form1.all_worker.value=""
end sub
</script>
<script language=vbscript>
sub menadd
	if document.form1.wk_item.value="" then
		document.form1.wk_item.value=document.form1.men_w.value
	else	
		document.form1.wk_item.value=document.form1.wk_item.value+" "+document.form1.men_w.value
	end if
end sub
sub dateadd
	if document.form1.wk_item.value="" then
		document.form1.wk_item.value=document.form1.date_w.value
	else	
		document.form1.wk_item.value=document.form1.wk_item.value+" "+document.form1.date_w.value
	end if
	document.form1.doing_date1.value=document.form1.date_w.value
end sub
sub timeadd
	if document.form1.wk_item.value="" then
		document.form1.wk_item.value=document.form1.time_w.value
	else	
		document.form1.wk_item.value=document.form1.wk_item.value+" "+document.form1.time_w.value
	end if
end sub
sub placeadd
	if document.form1.wk_item.value="" then
		document.form1.wk_item.value=document.form1.place_w.value
	else	
		document.form1.wk_item.value=document.form1.wk_item.value+" "+document.form1.place_w.value
	end if
end sub
sub thingadd
	if document.form1.wk_item.value="" then
		document.form1.wk_item.value=document.form1.thing_w.value
	else	
		document.form1.wk_item.value=document.form1.wk_item.value+" "+document.form1.thing_w.value
	end if
end sub
sub item_chk
	if document.form1.wk_item.value="" or document.form1.all_worker.value="" or document.form1.wk_class.value="" then
		ok=msgbox("請輸入工作分類、主旨及知會人員！！",0,"錯誤警告")
	else
	end if
end sub
sub press_chk
	if document.form1.wk_item.value="" or document.form1.all_worker.value="" or document.form1.wk_class.value="" then
		ok=msgbox("請輸入工作分類、主旨及知會人員！！",0,"錯誤警告")
	else
	     form1.submit
	end if
end sub
</script>
<font size=4 color="red"><b><%=worker%>重新工作公告單(原公告編號為<%=wk_id%>)</b></font>
<table border=1 cellspacing=0 cellpadding=0>
<col width=100>
<col width=130>
<col width=100>
<col width=130>
<col width=100>
<col width=130>
<tr>
	<td align="right">公告者：</td>
	<td><input type='text' name='wk_order' value='<%=worker%>' style="width:100%;" readonly></td>
	<td align="right">公告日期：</td>
	<td><input type='text' name='undo_date1' value='<%=undo_date1%>' style="width:100%;" readonly></td>
	<td align="right"><font color="red">工作分類：</font></td>
	<td>
	<input type='text' name='wk_class' value='A' style="width:100%;">
	</td>
</tr>
<tr>
	<td align="right">
	<font color="red">主旨：</font>
	</td>
	<td colspan=3>
	<input type='text' name='wk_item' value='<%=wk_item%>(原公告編號為<%=wk_id%>)' style="width:100%;" onchange="item_chk">
	</td>
	<td align="right"><font color="red">執行日期：</font></td>
	<td><input type='text' name='doing_date1' value='<%=date()%>' style="width:100%;"></td>

</tr>
<tr>
	<td colspan=6 align="center">
		<!-- #Include file = "./include/workitem.inc" -->
	</td>
</tr>
<tr>
	<td align="right">
	<font color="red">執行內容：</font>
	</td>
	<td colspan=5>
	<TEXTAREA name="wk_content" rows="5" style="width:100%;" ><%=wk_content%></TEXTAREA>
	</td>
</tr>
<tr>
	<td align="right">
	<font color="red">知會人員：</font>
	</td>
	<td colspan=5>
	<input type='text' name='all_worker' value='<%=wk_checker%>' size='85' onchange="item_chk">
	</td>
</tr>
<tr>
	<td colspan=6 align="center">
<%	
	for i=1 to worker_no
%>
	<input type="button"  name="worker_s<%=i%>" value="<%=worker_a(i-1)%>" onclick="worker_s<%=i%>_click">
<%
	next
%>	
	<!-- <input type="button" name="all_sel" value="全部人員" onclick="all_sel_click"> -->
	<input type="button" name="all_unsel" value="清除人員" onclick="all_unsel_click">
	</td>
</tr>
<tr>
	<td colspan=6 align="center">
	<input type="button" name="press" value="確定公告" onclick="press_chk">
<!-- 		<input type="submit" name="press" value="確定公告" onmouseover="press_chk">
 -->	<input type="reset" name="cancel" value="清除資料" disabled>
	</td>
<tr>
</table>

</form>
<center>
</body>
</html>
