<%@ Language=VBScript CODEPAGE=950 %>
<%
	'Ū���H���m�W
	worker = Session("worker")
	wk_id=Request("wk_id")
'	p_wk_content=trim(request("wk_content"))
'	p_wk_item=trim(request("wk_item"))
'	p_doing_date1=request("doing_date1")
	p_wk_group="�M�פu�@"
%>
<html>
<head>
<title>��ƭק�</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'�з���';background-color:'##FFFFcc'}
--></style>
</head>
<body>
<center>
<!-- �}�Ҹ�Ʈw -->
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
	'----------2003/03/15�ץ� 
	'�ק��Ƥ�SQL���O�r�� �������
	'strSQL_edit="Update "&tb_name&" set wk_content='"&request("wk_content")&"'"
	'strSQL_edit=strSQL_edit & ",doing_date1=#"& request("doing_date1") &"#"
	'strSQL_edit=strSQL_edit & ",wk_item='"& request("wk_item") &"'"
	'strSQL_edit=strSQL_edit & " where wk_id =" & wk_id
	'����SQL���O
	'conDB.Execute strSQL_edit
	'---------------------------------------------------------
'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_id="&wk_id
rstObj1.open strSQL_show,conDB,1,3
rstobj1.MoveFirst
'Ū�����
rstObj1.fields("wk_content")= date()&"�ର�M�פu�@�G"&chr(13)&rstObj1.fields("wk_content")
'rstObj1.fields("doing_date1")= p_doing_date1
'rstObj1.fields("wk_item")= p_wk_item
rstObj1.fields("wk_group")= p_wk_group
rstObj1.UpdateBatch
'������ƶ�
rstObj1.Close
'���]����ܼ� 
set rstObj1=Nothing
'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing 

strURL1="wk_show.asp?wk_id="&wk_id
response.redirect(strURL1)
%>

</body>
</html>
