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
firstlogin=Request("add_chk")
if Request("add_chk")<>"add" then
%>
<script language=vbscript>
function form1_onSubmit
	dim worker_arr(<%=worker_no%>)
	dim worker_no
	worker_no=<%=worker_no%>
	
	add_name=Trim(document.form1.add_name.value)
	if add_name="" then
		ok=msgbox("請輸入姓名!",0,"錯誤警告")
		form1_onSubmit=false
	else
	<%	for i=1 to worker_no 
	%>
			worker_arr(<%=i-1%>)="<%=worker_a(i-1)%>"
	<%	next
	%>
		for i=1 to worker_no
			if add_name=worker_arr(i-1) then
				msgbox "有相同人員之資料存在！！",0,"同名警告"
				form1_onSubmit=false
				exit for
			else
				form1_onSubmit=true
			end if
		next		
	end if
end function
</script>

	<font face="標楷體" size=5><b>新增工作人員畫面</b></font><br>
	<FORM name="form1" action="wkr_add.asp" method=post>
	<input type=hidden name="add_chk" value="add">
		<table>
		<tr><td>暱稱 :</td>
		<td><INPUT type=text name="add_name"></td></tr>
		<tr><td>性別 :</td>
		<td><Input type="radio" name="add_sex" value="男" checked>男
		    <Input type="radio" name="add_sex" value="女">女
		</td></tr>
		<tr><td>英文名字 :</td>
		<td><INPUT type=text name="add_name_e"></td></tr>
		</table>
		<INPUT type=submit name="press" value=送出資料>
		<INPUT type=reset name="cancel" value=清除輸入>
	</FORM>
	<font face="標楷體" size=4><A Href='wkr_del.asp?del_chk=ok'>刪除工作人員</A></font>&nbsp;&nbsp;
	<font face="標楷體" size=4><A Href='firstpage.asp'>回首頁</A></font>
<%
else
%>
	<%
	add_name=Request("add_name")
	add_sex=Request("add_sex")
	add_name_e=Request("add_name_e")
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
	'建立SQL字串,新增資料列到資料表worker中	
	strSQL="insert into " & tb_name_a &"(worker,wkr_sex,wkr_pwd,e_name) Values('"
	strSQL=strSQL & add_name & "','" & add_sex & "','6980','"& add_name_e &"')"
	'執行SQL指令
	conDB_a.Execute StrSQL
	'關閉資料庫 
	conDB_a.Close
	'重設物件變數 
	set conDB_a=Nothing 
	%>	
		<table>
		<tr><td>暱稱 :</td><td><%=add_name%></td></tr>
		<tr><td>性別 :</td><td><%=add_sex%></td></tr>
		<tr><td>英文名字 :</td><td><%=add_name_e%></td></tr>
		</table>
	<p><font face="標楷體" size=4>新工作人員資料新增完成</font></p>
	<font face="標楷體" size=3><A Href='wkr_add.asp?add_chk=ok'>新增工作人員</A></font>&nbsp;&nbsp;
	<font face="標楷體" size=4><A Href='wkr_del.asp?del_chk=ok'>刪除工作人員</A></font>&nbsp;&nbsp;
	<font face="標楷體" size=3><A Href='firstpage.asp'>回首頁</A></font>
<%
end if
%>
</center>
</body>
</html>
