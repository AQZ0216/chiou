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
firstlogin=Request("del_chk")
if Request("del_chk")<>"del" then

%>
	<font face="�з���" size=5><b>�R���u�@�H���e��</b></font><br>
	<FORM name="form1" action="wkr_del.asp" method=post>
	<input type=hidden name="del_chk" value="del">
		<table>
		<tr><td>�ʺ� :</td>
		<td>
		<SELECT name="del_name" size=1 >
		<option value=""></option>
	<%
	'�]�w�U�ԥ\��m�W�ƶ� 
	for i=1 to worker_no
		IF worker_a(i-1)=worker then
			Response.Write("<OPTION value=" & worker_a(i-1) & " selected ><font face='�з���' size=5>" & worker_a(i-1)&"</font></OPTION>")
		Else
			Response.Write("<OPTION value=" & worker_a(i-1) & "><font face='�з���' size=5>" & worker_a(i-1)&"</font></OPTION>")	
		End IF
	next
	%>
	
		</SELECT>
		</td></tr>
		</table>
		<INPUT type=submit name="press" value=�R�����>
		<INPUT type=reset name="cancel" value=�M����J>
	</FORM>
	<font face="�з���" size=4><A Href='wkr_add.asp?add_chk=ok'>�s�W�u�@�H��</A></font>&nbsp;&nbsp;
	<font face="�з���" size=4><A Href='firstpage.asp'>�^����</A></font>
<%
else
%>
	<%
	del_name=Request("del_name")
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
	'�إ�SQL�r��
	strSQL="Delete From " & tb_name_a &" Where worker = '"& del_name &"'"		
	'����SQL���O
	conDB_a.Execute StrSQL
	'������Ʈw 
	conDB_a.Close
	'���]�����ܼ� 
	set conDB_a=Nothing 
	%>
<p><font face="�з���" size=4>�u�@�H��</font><font face="�з���" size=5>[<%=del_name%>]</font><font face="�з���" size=4>��Ƥw�g�R��</font></p>	

	<font face="�з���" size=3><A Href='wkr_add.asp?add_chk=ok'>�s�W�u�@�H��</A></font>&nbsp;&nbsp;
	<font face="�з���" size=4><A Href='wkr_del.asp?del_chk=ok'>�R���u�@�H��</A></font>&nbsp;&nbsp;
	<font face="�з���" size=3><A Href='firstpage.asp'>�^����</A></font>
<%
end if
%>

</center>
</body>
</html>
