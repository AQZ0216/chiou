<% @codepage=950%>
<!-- #Include file = "./include/array_worker.inc" -->
<html>
<HEAD>
<title>工作管理系統</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</HEAD>
<body bgcolor="#F0FFF0">
<center>
<%
firstlogin=Request("del_chk")
if Request("del_chk")<>"del" then

%>
	<font face="標楷體" size=5><b>刪除工作人員畫面</b></font><br>
	<FORM name="form1" action="wkr_del.asp" method=post>
	<input type=hidden name="del_chk" value="del">
		<table>
		<tr><td>暱稱 :</td>
		<td>
		<SELECT name="del_name" size=1 >
		<option value=""></option>
	<%
	'設定下拉功能姓名事項 
	for i=1 to worker_no
		IF worker_a(i-1)=worker then
			Response.Write("<OPTION value=" & worker_a(i-1) & " selected ><font face='標楷體' size=5>" & worker_a(i-1)&"</font></OPTION>")
		Else
			Response.Write("<OPTION value=" & worker_a(i-1) & "><font face='標楷體' size=5>" & worker_a(i-1)&"</font></OPTION>")	
		End IF
	next
	%>
	
		</SELECT>
		</td></tr>
		</table>
		<INPUT type=submit name="press" value=刪除資料>
		<INPUT type=reset name="cancel" value=清除輸入>
	</FORM>
	<font face="標楷體" size=4><A Href='wkr_add.asp?add_chk=ok'>新增工作人員</A></font>&nbsp;&nbsp;
	<font face="標楷體" size=4><A Href='firstpage.asp'>回首頁</A></font>
<%
else
%>
	<%
	del_name=Request("del_name")
	'開啟資料庫
	' 連結Access資料庫daywork.mdb
	DBpath_a=Server.MapPath("./database/daywork.mdb")
	strCon_a="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_a
	'建立資料庫連結物件
	set conDB_a= Server.CreateObject("ADODB.Connection")
	'連結資料庫	
	conDB_a.Open strCon_a
	'開啟資料表名稱
	tb_name_a="worker_data"	
	'建立SQL字串
	strSQL="Delete From " & tb_name_a &" Where worker = '"& del_name &"'"		
	'執行SQL指令
	conDB_a.Execute StrSQL
	'關閉資料庫 
	conDB_a.Close
	'重設物件變數 
	set conDB_a=Nothing 
	%>
<p><font face="標楷體" size=4>工作人員</font><font face="標楷體" size=5>[<%=del_name%>]</font><font face="標楷體" size=4>資料已經刪除</font></p>	

	<font face="標楷體" size=3><A Href='wkr_add.asp?add_chk=ok'>新增工作人員</A></font>&nbsp;&nbsp;
	<font face="標楷體" size=4><A Href='wkr_del.asp?del_chk=ok'>刪除工作人員</A></font>&nbsp;&nbsp;
	<font face="標楷體" size=3><A Href='firstpage.asp'>回首頁</A></font>
<%
end if
%>

</center>
</body>
</html>
