<% @codepage=950%>
<%
errmsg=request("errmsg")
%>
<html>
<HEAD>
<title>工作管理系統</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</HEAD>
<body bgcolor="#F0FFF0">
<center>

	<FORM name="form1" action="wkr_add_login.asp" method=post>
		<table>
		<tr><td>密碼 :</td>
		<td><INPUT type="password" name="pwd01"></td>
		</tr>
		<tr>
		<td colspan=2>
		<INPUT type="submit" name="press" value="登入工作人員新增畫面" style="font-size:12pt">
		</td>
		</tr>
		</table>
<font style="color:red;letter-spacing:5pt;font-weight:bold;"><%=errmsg%></font>
<hr>
<font face="標楷體" size=4><A Href='firstpage_elink.asp'>修改網頁連結資料</A></font>
<hr>
<font face="標楷體" size=4><A Href='./firstpage.asp'>回工作管理系統</A></font>
<hr>
</FORM>
</center>
</body>
</html>
