<!-- �}�Ҹ�Ʈw -->
<%
'strbackURL=Session("strbackURL")
'�]�wŪ����ƽs��
	fw_id=Request("fw_id")
	wk_id=Request("wk_id")
' �s��Access��Ʈw./database/fileman.mdb
DBpath=Server.MapPath("./database/wk_file.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="wk_file"
%>
<%
'�R����Ƥ�SQL���O�r��
strSQL_del="Delete from " & tb_name & " where fw_id =" & fw_id
'����SQL���O
conDB.Execute strSQL_del
%>
<%
'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing 

str_url="wk_pj_show.asp?wk_id="&wk_id
response.redirect(str_url) 

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
<center>
</body>
</html>
