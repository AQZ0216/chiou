<% @codepage=950%>
<!-- #Include file = "./include/f_week_cstr.inc" -->
<%
	'讀取人員姓名
	worker = Session("worker")
	'讀取今天日期
	ckdate=date()+2
wkgroup="專案工作"
%>
<!-- 開啟資料庫 -->
<!-- #Include file = "./include/opendb_daywork.inc" -->

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

<%
'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
'strSQL_show="Select * from " & tb_name & " where wk_group like '%"&wkgroup&"%' and wk_doer like '%"&worker&"%' and doing_date1 <= #"&ckdate &"# and wk_undoer like '%"&worker&"%' order by doing_date1 desc"
'strSQL_show="Select * from " & tb_name & " where wk_group like '%"&wkgroup&"%' order by doing_date1 desc"
strSQL_show="Select * from " & tb_name & " where not(isnull(pj_02) or trim(pj_02) like '' or pj_02 like '基金會讀書會資料') or (wk_group like '%"&wkgroup&"%') order by doing_date1 desc"
rstObj1.open strSQL_show,conDB,3
totalput=rstObj1.recordcount
if totalput=0 then
%>
	<font size=4>無專案工作事項</font>
<%
else
%>
<table border=1 cellspacing=0 cellpadding=0>
<col width=50>
<col width=130>
<col width=420>
<col width=170>
<tr >
	<td colspan=4 align=center>
	<font size=4>所有專案工作事項共:<font color=red><%=totalput%></font>筆</font>
	</td>
</tr>
<tr >
	<td align=center>序號</td>
	<td align=center>執行日期</td>
<!--	<td align=center>公告者</td>-->
	<td align=center>主旨</td>
	<td align=center>
		<a href="./pj_add.asp" target="_blank"> <img src="./img/add1.gif" alt="專案新增" width="15" height="15" style='cursor:hand;border:0;'></a>
	專案名稱
		<a href="./pj_list.asp" target="_blank"> <img src="./img/list1.gif" alt="專案名稱列表" width="15" height="15" style='cursor:hand;border:0;'></a>
	</td>
</tr>
<%
	'列出資料項目
	rstobj1.MoveFirst
	for i=1 to totalput
	'讀取資料
		wk_id=rstObj1.fields("wk_id")
		undo_date1=rstObj1.fields("undo_date1")
		doing_date1=rstObj1.fields("doing_date1")
		wk_item=rstObj1.fields("wk_item")
		wk_order=rstObj1.fields("wk_order")
		pj_id=rstObj1.fields("pj_id")
		pj_02=rstObj1.fields("pj_02")
		Response.Write( "<tr>")		
		Response.Write( "<td align=center><font size=3>" & i &"</font></td>")		
		Response.Write( "<td align=center style='text-align:right;padding-right:2pt;'><font size=3>" & doing_date1 &" ("&week_cstr(doing_date1)&")</font></td>")
		'Response.Write( "<td align=center><font size=3>" & wk_order &"</font></td>")
		'Response.Write( "<td align=center><font size=3>" & wk_exe &"</font></td>")
		strA="<a href=wk_pj_show.asp?wk_id="& rstObj1.fields("wk_id")&">"
		Response.Write( "<td align=left>" & strA & wk_item &"</a></td>")

		if pj_id="" or isnull(pj_id) then
%>		
		<td align=left><font size=3>
	<a href="./pj_add.asp?wk_id=<%=wk_id%>" target="_blank"> <img src="./img/add1.gif" alt="新增專案名稱" width="15" height="15" style='cursor:hand;border:0;'></a>
	<a href="./pj_sel.asp?wk_id=<%=wk_id%>" target="_blank"> <img src="./img/sel1.gif" alt="選擇專案名稱" width="15" height="15" style='cursor:hand;border:0;'></a>
		</font></td>
<%
		else
%>		
		<td align=left><font size=3>
	<a href="./pj_delsel.asp?wk_id=<%=wk_id%>&p_id=<%=pj_id%>" target="_blank"> <img src="./img/del1.gif" alt="移除專案名稱" width="15" height="15" style='cursor:hand;border:0;'></a>
	<a href="./pj_show.asp?p_id=<%=pj_id%>" target="_blank"><%=pj_02%></a>
		</font></td>
<%
		end if
		Response.Write( "</tr>")
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
<center>
</body>
</html>
