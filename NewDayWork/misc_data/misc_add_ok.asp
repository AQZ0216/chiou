
<%
' �s��Access��Ʈwmisc_data.mdb
DBpath_b=Server.MapPath("./misc_data.mdb")
strCon_b="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_b
'�إ߸�Ʈw�s������
set conDB_b= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB_b.Open strCon_b
'Ū����ƪ�W��
s_id=(Request("add_item"))
'Ū�����W��
Select Case s_id
Case 1
	tb_name_b="writer_table"
	item="writer"
	iword="�H��"	
Case 2
	tb_name_b="place_table"
	item="place"
	iword="�a�I"
Case 3
	tb_name_b="thing_table"
	item="thing"
	iword="�ƥ�"
Case else
	response.redirect "../error/errormisc.htm"
End Select
if Request("keyword")="" then
	response.redirect "../error/errormisc.htm"
end if
'�s�W��Ƥ�SQL���O�r��
strSQL_add="Insert into "&tb_name_b&" ( "&item&" ) values ('" &Request("keyword")&"')"

'����SQL���O
conDB_b.Execute strSQL_add

'������Ʈw 
conDB_b.Close
'���]�����ܼ� 
set conDB_b=Nothing 
%>
<html>
<head>
<title>������Ʒs�W</title>
<style type="text/css"><!--
body{font-family:'�з���';background-color:'##FFFFcc'}
--></style>
</head>
<body >
<center>
<p><font size=5 color="#ff0000">�ﶵ�s�W�����I�I</font></p>
<p><font size=5 color=red>��<%=iword%>��</font><font size=5 color="#0000ff">�ﶵ���A�s�W</font><font size=5 color=red>��<%=Request("keyword")%>��</font><font size=5 color="#0000ff">����</font></p>
<table border=1><tr>
<td width=180 align=center><a href="./misc_edit.asp">�^�ﶵ�s�׭�</a></td>
</tr></table>
</center>
</body>
</html>