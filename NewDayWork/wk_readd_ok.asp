<%@ Language=VBScript CODEPAGE=950 %>
<!-- �}�Ҹ�Ʈw -->
<!-- #Include file = "./include/opendb_daywork.inc" -->
<%
'�P�_�O�_��J�u�@���� 
keyword=request("wk_class")
if keyword="" then 
	response.redirect("wk_add.asp")
else
end if

%>	
<html>
<head>
<title>��Ƨ���s�W</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'�з���';background-color:'##FFFFcc'}
--></style>
</head>
<body>
<center>

<!-- #Include file = "./include/readd_wk_ok.inc" -->

<%
'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing 
%>
<script language="Javascript">
	alert("��Ʒs�W�����I�I");
	location.href="wk_lst_undo.asp";
</script>

</body>
</html>
