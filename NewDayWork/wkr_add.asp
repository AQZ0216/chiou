<% @codepage=950%>
<!-- #Include file = "./include/array_worker.inc" -->
<html>
<HEAD>
<title>�u�@�޲z�t��</title>
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
		ok=msgbox("�п�J�m�W!",0,"���~ĵ�i")
		form1_onSubmit=false
	else
	<%	for i=1 to worker_no 
	%>
			worker_arr(<%=i-1%>)="<%=worker_a(i-1)%>"
	<%	next
	%>
		for i=1 to worker_no
			if add_name=worker_arr(i-1) then
				msgbox "���ۦP�H������Ʀs�b�I�I",0,"�P�Wĵ�i"
				form1_onSubmit=false
				exit for
			else
				form1_onSubmit=true
			end if
		next		
	end if
end function
</script>

	<font face="�з���" size=5><b>�s�W�u�@�H���e��</b></font><br>
	<FORM name="form1" action="wkr_add.asp" method=post>
	<input type=hidden name="add_chk" value="add">
		<table>
		<tr><td>�ʺ� :</td>
		<td><INPUT type=text name="add_name"></td></tr>
		<tr><td>�ʧO :</td>
		<td><Input type="radio" name="add_sex" value="�k" checked>�k
		    <Input type="radio" name="add_sex" value="�k">�k
		</td></tr>
		<tr><td>�^��W�r :</td>
		<td><INPUT type=text name="add_name_e"></td></tr>
		</table>
		<INPUT type=submit name="press" value=�e�X���>
		<INPUT type=reset name="cancel" value=�M����J>
	</FORM>
	<font face="�з���" size=4><A Href='wkr_del.asp?del_chk=ok'>�R���u�@�H��</A></font>&nbsp;&nbsp;
	<font face="�з���" size=4><A Href='firstpage.asp'>�^����</A></font>
<%
else
%>
	<%
	add_name=Request("add_name")
	add_sex=Request("add_sex")
	add_name_e=Request("add_name_e")
	'�}�Ҹ�Ʈw
	' �s��Access��Ʈwdaywork.mdb
	DBpath_a=Server.MapPath("./database/daywork.mdb")
	strCon_a="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_a
	'�إ߸�Ʈw�s������
	set conDB_a= Server.CreateObject("ADODB.Connection")
	'�s����Ʈw	
	conDB_a.Open strCon_a
	'�}�Ҹ�ƪ�W��
	tb_name_a="worker_data"	
	'�إ�SQL�r��,�s�W��ƦC���ƪ�worker��	
	strSQL="insert into " & tb_name_a &"(worker,wkr_sex,wkr_pwd,e_name) Values('"
	strSQL=strSQL & add_name & "','" & add_sex & "','6980','"& add_name_e &"')"
	'����SQL���O
	conDB_a.Execute StrSQL
	'������Ʈw 
	conDB_a.Close
	'���]�����ܼ� 
	set conDB_a=Nothing 
	%>	
		<table>
		<tr><td>�ʺ� :</td><td><%=add_name%></td></tr>
		<tr><td>�ʧO :</td><td><%=add_sex%></td></tr>
		<tr><td>�^��W�r :</td><td><%=add_name_e%></td></tr>
		</table>
	<p><font face="�з���" size=4>�s�u�@�H����Ʒs�W����</font></p>
	<font face="�з���" size=3><A Href='wkr_add.asp?add_chk=ok'>�s�W�u�@�H��</A></font>&nbsp;&nbsp;
	<font face="�з���" size=4><A Href='wkr_del.asp?del_chk=ok'>�R���u�@�H��</A></font>&nbsp;&nbsp;
	<font face="�з���" size=3><A Href='firstpage.asp'>�^����</A></font>
<%
end if
%>
</center>
</body>
</html>
