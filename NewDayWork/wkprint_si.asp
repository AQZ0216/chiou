<%@ Language=VBScript CODEPAGE=950 %>
<%
	'Ū���H���m�W
	worker = Session("worker")
	wk_id=Request("wk_id")
%>
<!-- �}�Ҹ�Ʈw -->
<!-- #Include file = "./include/opendb_daywork.inc" -->
<%
'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_id =" & wk_id
rstObj1.open strSQL_show,conDB,3,1
'Ū�����
undo_date1=rstObj1.fields("undo_date1")
doing_date1=rstObj1.fields("doing_date1")
done_date1=rstObj1.fields("done_date1")
wk_item=rstObj1.fields("wk_item")
wk_content=rstObj1.fields("wk_content")
wk_order=rstObj1.fields("wk_order")
wk_doer=rstObj1.fields("wk_doer")
wk_checker=rstObj1.fields("wk_checker")
wk_undoer=rstObj1.fields("wk_undoer")
wk_class=rstObj1.fields("wk_class")
%>
<%
'������ƶ�
rstObj1.Close
'���]����ܼ� 
set rstObj1=Nothing
'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing 

'�N����r��chr(13)�নhtml�����������<br>
if isnull(wk_content) then
	wk_contenta="(�ť�)"
else
	wk_contenta=replace(wk_content,chr(13),"<br>")
end if

%>

<html>
<head>
<title>�u�@�C�L</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="icon" href="../daywork/img/khouse.ico" type="image/ico" />
<style type="text/css"><!--
body{font-family:'�s�ө���';background-color :'#FFFEEE'}

td{
	margin:5px 0 0 0;		/*��t�W�U���k*/
	border-color:'#000000'; /*���~���C��*/ 
	border-style:solid;		/*���~�ؽu��*/
	border-width:1px;		/*���~�ثp��*/  
	vertical-align:middle;	/*�r�髫������覡*/
	}
table{	
	/*margin:0 0 0 0;		/*��t�W�U���k*/
	border-collapse:collapse; 	/*��اΦ����X*/
	}
--></style>
</head>
<body>
<center>
<font size=5 color='blue'>�u�@�C�L</font><br>
<table border=0 cellspacing=0 cellpadding=2 style="width:600px" >
<col style="width:16%;background-color:#d3d3d3;" align=right>
<col style="width:16%;" align=left>
<col style="width:16%;background-color:#d3d3d3;" align=right>
<col style="width:16%;" align=left>
<col style="width:16%;background-color:#d3d3d3;" align=right>
<col style="width:16%;" align=left>
<tr>
	<td>�u�@�s���G
	<td><%=wk_id%>
	<td>���i����G
	<td><%=undo_date1%>
	<td>�������G
	<td><%=doing_date1%>
<tr>
	<td>�D���G</td>
	<td colspan=5 ><%=wk_item%>
<tr>
	<td colspan=6 style="text-align:center;background-color:#d3d3d3;">���椺�e�G
<tr>
	<td colspan=6 style="text-align:left;background-color:#FFFEEE;padding: 6px 6px 6px 12px;" >
	
	<%=wk_contenta%><br>
	
</table>
</center>

</body>

</html>
