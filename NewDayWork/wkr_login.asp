<% @codepage=950%>
<%
errmsg=request("errmsg")
%>
<html>
<HEAD>
<title>�u�@�޲z�t��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</HEAD>
<body bgcolor="#F0FFF0">
<center>

	<FORM name="form1" action="wkr_add_login.asp" method=post>
		<table>
		<tr><td>�K�X :</td>
		<td><INPUT type="password" name="pwd01"></td>
		</tr>
		<tr>
		<td colspan=2>
		<INPUT type="submit" name="press" value="�n�J�u�@�H���s�W�e��" style="font-size:12pt">
		</td>
		</tr>
		</table>
<font style="color:red;letter-spacing:5pt;font-weight:bold;"><%=errmsg%></font>
<hr>
<font face="�з���" size=4><A Href='firstpage_elink.asp'>�ק�����s�����</A></font>
<hr>
<font face="�з���" size=4><A Href='./firstpage.asp'>�^�u�@�޲z�t��</A></font>
<hr>
</FORM>
</center>
</body>
</html>
