<%@ Language=VBScript CODEPAGE=950 %>
<%
	'Ū���H���m�W
	worker = Session("worker")
	wk_id=Request("wk_id")
%>
<!-- �}�Ҹ�Ʈw -->
<!-- Include file = "./include/opendb_daywork.inc" -->

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

' �s��Access��Ʈwdaywork.mdb
DBpath=Server.MapPath("./database/daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="work_data"

'�R����Ƥ�SQL���O�r��
strSQL_del="Delete from " & tb_name & " where wk_id =" & wk_id
'����SQL���O
conDB.Execute strSQL_del

'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing 
%>

<%

' �s��Access��Ʈwdaywork.mdb
DBpath=Server.MapPath("./database/temp-daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="work_data"

'�R����Ƥ�SQL���O�r��
strSQL_del="Delete from " & tb_name & " where (tmp_id =" & wk_id&" and ipt_ok=0)"
'����SQL���O
conDB.Execute strSQL_del

'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing 
%>
<script language="Javascript">
	alert("��ƧR�������I�I");
	location.href="wk_Calendar_all.asp";
</script>
</body>
</html>
