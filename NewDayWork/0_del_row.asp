<!-- �}�Ҹ�Ʈw -->
<%
'�]�wŪ����ƽs��
last_row=Request("last_row")    'p_id
'=======================================================================
' �s��Access��Ʈw./database/linkweb.mdb
DBpath=Server.MapPath("./database/linkweb.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="linkdata"
'�R����Ƥ�SQL���O�r��
strSQL_del="Delete from " & tb_name & " where lk_row =" & last_row
'����SQL���O
conDB.Execute strSQL_del
'������Ʈw
conDB.Close
'���]�����ܼ� 
set conDB=Nothing
'=======================================================================

nexturl="firstpage_elink.asp"
response.redirect(nexturl)
%>	
<html>
<head>
<title>�T�w�R�����</title>
<style type="text/css"><!--
body{font-family:'�s�ө���';background-color :'#FFFEEE'}
input{font-family:'�s�ө���';font-size:12pt;cursor:hand;}
select{font-family:'�s�ө���';font-size:10pt;cursor:hand;}
--></style>
</head>
<body>
<center></center>
</body>
</html>