<% @codepage=950%>
<%
wkgroup="一般工作"
	'讀取人員姓名
	worker = Session("worker")
	'讀取今天日期
	q_key31=request("q_key31")'日期型式 
	q_key32=request("q_key32")'判斷式
	q_key33=request("q_key33")'查詢日期
	
Select Case q_key31
    Case "完成日期" 
        q_key31_p="done_date1"
    Case "執行日期" 
		q_key31_p="doing_date1"
    Case "派工日期" 
		q_key31_p="undo_date1"
    Case Else

End Select

Select Case q_key32
    Case "等於" 
        q_key32_p="="
    Case "大於等於" 
        q_key32_p=">="
    Case "小於等於" 
        q_key32_p="<="
    Case "大於" 
        q_key32_p=">"
    Case "小於" 
        q_key32_p="<"
    Case Else

End Select
q_key33_p=cdate(q_key33)	
%>
<!-- 開啟資料庫 -->
<!-- #Include file = "./include/opendb_daywork.inc" -->

<HTML>
<HEAD>
<title>工作管理系統</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'微軟正黑體';background-color:'#F0FFF0'}
input{font-family:'微軟正黑體';font-size:12pt;}
textarea{font-family:'微軟正黑體';}
SELECT{font-family:'微軟正黑體';font-size:12pt;}
td{font-family:'微軟正黑體';}
.tdtext{
		font-size:4mm;
		} 
.tittext{
		font-size:4.5mm;
		font-weight:bold;
		} 
--></style>
</HEAD>
<BODY>
<center>
<font style="font-size:5mm;color:#0000ff;">
查詢條件：[<font color=red><%=q_key31%></font>][<font color=red><%=q_key32%></font>][<font color=red><%=q_key33%></font>] 
</font>
<%
'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_group like '%"&wkgroup&"%' and "&q_key31_p&" "&q_key32_p&" #"&q_key33_p&"# order by doing_date1 desc"
rstObj1.open strSQL_show,conDB,3
totalput=rstObj1.recordcount
if totalput=0 then
%>
	<font size=4>查詢結果：無工作事項</font>
<%
else
%>
<table border=1 cellspacing=0 cellpadding=0>
<col width=50>
<col width=100>
<col width=420>
<col width=200>
<tr >
	<td colspan=4 align=center>
	<font size=4>查詢結果：工作事項共有<font color=red><%=totalput%></font>筆</font>
	</td>
</tr>
<tr >
	<td align=center>序號</td>
	<td align=center>執行日期</td>
	<td align=center>主旨</td>
	<td align=center>
		<a href="./pj_add.asp" target="_blank"> <img src="./img/add1.gif" alt="專案新增" width="15" height="15" style='cursor:hand;border:0;'></a>
	專案名稱
		<a href="./pj_list.asp" target="_blank"> <img src="./img/list1.gif" alt="專案列表" width="15" height="15" style='cursor:hand;border:0;'></a>
	</td>
</tr>
<%
	'列出資料項目
	rstobj1.MoveFirst
	for i=1 to totalput
	'讀取資料
		wk_gp=trim(rstObj1.fields("wk_group"))
		wk_id=rstObj1.fields("wk_id")
		undo_date1=rstObj1.fields("undo_date1")
		doing_date1=rstObj1.fields("doing_date1")
		wk_item=rstObj1.fields("wk_item")
		wk_order=rstObj1.fields("wk_order")
		pj_id=rstObj1.fields("pj_id")
		pj_02=rstObj1.fields("pj_02")
		Response.Write( "<tr>")
		Response.Write( "<td align=center><font size=3>" & i &"</font></td>")		
		Response.Write( "<td align=center><font size=3>" & doing_date1 &"</font></td>")
		'Response.Write( "<td align=center><font size=3>" & wk_order &"</font></td>")
		strA="<a href=wk_show.asp?wk_id="& rstObj1.fields("wk_id")&">"
		if wk_gp="一般工作" then
			Response.Write( "<td align=left>" & strA & wk_item &"</a></td>")
		else
			strA1="<<專案工作>>"
			Response.Write( "<td align=left style='background-color:#ffff99;'>" & strA &strA1 & wk_item &"</a></td>")
		end if
		
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
