<%@ Language=VBScript CODEPAGE=950 %>
<%
'�ഫ����榡���
Function FormatSQLDate(dDate)
	Select Case Vartype(dDate)
	Case 7: '����ɶ�
		FormatSQLDate="#"&FormatDateTime(dDate,2)&"#"
	Case 8: '�r��
		If IsDate(dDate) then
			FormatSQLDate="#"&FormatDateTime(dDate,2)&"#"
		Else
			FormatSQLDate="NULL"
		End if
	Case Else
		FormatSQLDate="NULL"
	End select
End function

%>
<!-- �}�Ҹ�Ʈw -->
<%

'�]�wŪ����ƽs��
new_row=Request("new_row")    'p_id

' �s��Access��Ʈw./database/linkweb.mdb
DBpath=Server.MapPath("./database/linkweb.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="linkdata"
'======================================================================
for zd=1 to 6

'�s�W��Ƥ�SQL���O�r��
strSQL_add="Insert into "& tb_name &" (lk_row,lk_col,lk_show) values ('"
strSQL_add=strSQL_add & new_row &"','"
strSQL_add=strSQL_add & zd &"',"

strSQL_add=strSQL_add & true &")"
'����SQL���O
conDB.Execute strSQL_add
next
'======================================================================

'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing
 
nexturl="firstpage_elink.asp"
response.redirect(nexturl)
%>

	
<html>
<head>
<title>�T�w�s�W���</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'�s�ө���';background-color :'#FFFEEE'}
input{font-family:'�s�ө���';font-size:12pt;cursor:hand;}
select{font-family:'�s�ө���';font-size:10pt;cursor:hand;}
--></style>
</head>
<body>
</body>
</html>
