<%@ Language=VBScript CODEPAGE=950 %>
<%
	'Ū���H���m�W
	worker = Session("worker")
	wk_id=Request("wk_id")
%>
<!-- �}�Ҹ�Ʈw -->
<!-- #Include file = "./include/opendb_daywork.inc" -->

<html>
<head>
<title>�T�w�R�����</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'�з���';background-color:'##FFFFcc'}
--></style>
</head>
<body>
<center>
<%
'�R����Ƥ�SQL���O�r��
strSQL_del="Delete from " & tb_name & " where wk_id =" & wk_id
'����SQL���O
conDB.Execute strSQL_del
%>
<%
'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing

   strbackURL=session("strbackURL")
   if strbackURL="" or isnull(strbackURL) then strbackURL="wk_calendar_all.asp"
   response.redirect(strbackURL)
%>
<!-- <script language="Javascript">
	alert("��ƧR�������I�I");
	location.href="wk_lst_doing.asp";
</script> -->
</body>
</html>
