<%@ Language=VBScript CODEPAGE=950 %>
<!-- �}�Ҹ�Ʈw -->
<%
' �s��Access��Ʈwdaywork.mdb
DBpath=Server.MapPath("./database/wk_file.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="wk_file"
'
wk_id=request("wk_id") 
fl_name=request("filename")

'�s�W��Ƥ�SQL���O�r��
strSQL_add="Insert into "&tb_name&" ("
strSQL_add=strSQL_add & "wk_id,"
strSQL_add=strSQL_add & "fl_name) values ('"
strSQL_add=strSQL_add & wk_id &"','"
strSQL_add=strSQL_add & fl_name &"')"

'����SQL���O
conDB.Execute strSQL_add


'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " order by fw_id desc"
rstObj1.open strSQL_show,conDB,3,3
totalput=rstObj1.recordcount
	'�C�X��ƶ���
rstobj1.MoveFirst
fw_id1=rstObj1.fields("fw_id")

'������ƶ�
rstObj1.Close
'���]����ܼ� 
set rstObj1=Nothing
'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing 

'str_url="flwk_display.asp?fw_id="&fw_id1

str_url="wk_pj_show.asp?wk_id="&wk_id
response.redirect(str_url) 


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

</body>
</html>
