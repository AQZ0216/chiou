<%@ Language=VBScript CODEPAGE=950 %>
<%
botton_color="#c3c3c3"
%>
<%
pj_id=request("p_id")
'pj_02=request("pj_02")

if pj_id="" or isnull(pj_id) then ckp1=1
if ckp1=1 then Response.redirect "pj_list.asp"
'if pj_02="" or isnull(pj_02) then ckp2=1
'if ckp1=1 and ckp2=1 then Response.redirect "pj_list.asp"

%>
<%
'�]�wŪ����ƽs��
'if pj_id="" or isnull(pj_id) then
'	Session("flstrbackURL")="pj_show.asp?pj_02="&pj_02
'else
	Session("flstrbackURL")="pj_show.asp?p_id="&pj_id
'end if

'�s��Access��Ʈw./database/daywork.mdb
DBpath=Server.MapPath("./database/daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="project_data"

%>
<%
'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
'if pj_id="" or isnull(pj_id) then
'	strSQL_show="Select * from " & tb_name & " where pj_02 like '"& pj_02 &"'"
'else
	strSQL_show="Select * from " & tb_name & " where pj_id =" & pj_id
'end if
rstObj1.open strSQL_show,conDB,3,1

p_00=rstObj1.fields("pj_id")	'�M��id
p_01=rstObj1.fields("pj_01")	'�M�׽s��
p_02=rstObj1.fields("pj_02")	'�M�צW��
p_03=rstObj1.fields("pj_03")	'�M�׻���

'������ƶ�
rstObj1.Close
'���]����ܼ� 
set rstObj1=Nothing
'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing 

botton_color="#c3c3c3"
%>	
<html>
<head>
<title>�R���M�׸��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'�s�ө���';background-color :'#FFFEEE'}
input{
	font-family:'�s�ө���';
	font-size:12pt;
	}
select{font-family:'�s�ө���';font-size:10pt;cursor:hand;}
.itxt{
	font-family:'�s�ө���';
	font-size:12pt;
	width:100%;
	height:100%;
	}
input.imenu { 
	/*font-size:15px;				/*�r��j�p*/
	/*font-weight:bold;
	cursor:hand;				/*��ЧΦ�*/ 
	background-color:'<%=botton_color%>'; 		
	margin:0 0 0 0;		/*��t�W�U���k*/
	width:100px;
	/*height:100%;*/
	color:#000000;
	letter-spacing:2px;
	cursor:hand;
     }
td{
	margin:0 0 0 0;		/*��t�W�U���k*/
	border-color:'#808080'; /*���~���C��*/ 
	border-style:solid;		/*���~�ؽu��*/
	border-width:1px;		/*���~�ثp��*/  
	vertical-align:middle;	/*�r�髫������覡*/
	font-size:15px;
	}
table{	
	margin:0 0 0 0;		/*��t�W�U���k*/
	border-collapse:collapse; 	/*��اΦ����X*/
	}
input.itext { 
	font-size:3.5mm;				/*�r��j�p*/
	/*cursor:hand;				/*��ЧΦ�*/ 
	width:100%;
	height:5mm;
	background-color:'#ffeedd'; 		/*�~���C��*/ 
	margin:0 0 0 0;		/*��t�W�U���k*/
	color:black;
     }

--></style>
</head>
<body>
<center>
<form name="form1" method=post action="pj_del_ok.asp">
<input type=hidden name="p00" value="<%=p_00%>">
<!--<input type=hidden name="p02" value="<%=p_02%>">-->
<font style="font-size:20px;font-weight:bold;letter-spacing:15px;">�i�M�צW�١j�R��</font><br>
<table border=0 cellspacing=0 cellpadding=2 style="width:760px" >
<col style="width:100px;" align=right>
<col style="" align=left>
<tr style="height:25px;">
<td colspan=2 align='center'>
<font style="color:red;font-size:5mm;font-weight:bold;">
�O�_�T�w�n�R���M�צW�١H&nbsp;&nbsp; 
</font>
	<input class=imenu type=submit name="sent" value="�T�w�R��"  onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='<%=botton_color%>';" >&nbsp;
	<input class=imenu type=button name=giveup value="�^�W�@��" onclick="history.back()"  onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='<%=botton_color%>';">
</td></tr>
<tr>
<td><font color="red">�M�׽s���G </font></td>
<td>
<%=p_01%>
</td>
</tr>
<tr>
<td><font color="red">�M�צW�١G </font></td>
<td>
<%=p_02%>
</td>
</tr>
<tr>
<td><font color="red">�M�׻����G </font></td>
<td>
<%=p_03%>
</td>
</tr>
</table>

<!--��ܱM�׻P����u�@-->
<%
'�s��Access��Ʈw./database/daywork.mdb
DBpath=Server.MapPath("./database/daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="work_data"
%>
<%
wkgroup="�M�פu�@"
'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_group like '%"&wkgroup&"%' and pj_id ="& pj_id &" order by doing_date1 desc"
rstObj1.open strSQL_show,conDB,3
totalput=rstObj1.recordcount
if totalput=0 then
%>
	<font size=4>�L�M�פu�@�ƶ�</font>
<%
else
%>
<table border=0 cellspacing=0 cellpadding=2 style="width:760px" >
<col style="width:50px;" align=center>
<col style="width:100px;" align=center>
<col style="" align=left>
<tr style="height:25px;">
	<td colspan=3 align=center>
	<font size=4>�Ҧ��M�פu�@�ƶ��@:<font color=red><%=totalput%></font>��</font>
	</td>
</tr>
<tr >
	<td align=center>�Ǹ�</td>
	<td align=center>������</td>
	<td align=center>�D��</td>
	</td>
</tr>
<%
	'�C�X��ƶ���
	rstobj1.MoveFirst
	for i=1 to totalput
	'Ū�����
		wk_id=rstObj1.fields("wk_id")
		undo_date1=rstObj1.fields("undo_date1")
		doing_date1=rstObj1.fields("doing_date1")
		wk_item=rstObj1.fields("wk_item")
		'pj_02=rstObj1.fields("pj_02")
		Response.Write( "<tr>")		
		Response.Write( "<td align=center><font size=3>" & i &"</font></td>")		
		Response.Write( "<td align=center><font size=3>" & doing_date1 &"</font></td>")
		strA="<a href=wk_pj_show.asp?wk_id="& rstObj1.fields("wk_id")&">"
		Response.Write( "<td align=left>" & strA & wk_item &"</a></td>")
		Response.Write( "</tr>")
	'����U�@���O��
		rstObj1.MoveNext
		if rstObj1.EOF=True then exit for
	next	
%>
</table>
<%
end if
'������ƶ�
rstObj1.Close
'���]����ܼ� 
set rstObj1=Nothing
%>

<%
wkgroup="�@��u�@"
'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_group like '%"&wkgroup&"%' and pj_id ="& pj_id &" order by doing_date1 desc"
rstObj1.open strSQL_show,conDB,3
totalput=rstObj1.recordcount
if totalput=0 then
%>
	<font size=4>�L�@��u�@�ƶ�</font>
<%
else
%>
<table border=0 cellspacing=0 cellpadding=2 style="width:760px" >
<col style="width:50px;" align=center>
<col style="width:100px;" align=center>
<col style="" align=left>
<tr style="height:25px;">
	<td colspan=3 align=center>
	<font size=4>�Ҧ��@��u�@�ƶ��@:<font color=red><%=totalput%></font>��</font>
	</td>
</tr>
<tr >
	<td align=center>�Ǹ�</td>
	<td align=center>������</td>
	<td align=center>�D��</td>
	</td>
</tr>
<%
	'�C�X��ƶ���
	rstobj1.MoveFirst
	for i=1 to totalput
	'Ū�����
		wk_id=rstObj1.fields("wk_id")
		undo_date1=rstObj1.fields("undo_date1")
		doing_date1=rstObj1.fields("doing_date1")
		wk_item=rstObj1.fields("wk_item")
		'pj_02=rstObj1.fields("pj_02")
		Response.Write( "<tr>")		
		Response.Write( "<td align=center><font size=3>" & i &"</font></td>")		
		Response.Write( "<td align=center><font size=3>" & doing_date1 &"</font></td>")
		strA="<a href=wk_pj_show.asp?wk_id="& rstObj1.fields("wk_id")&">"
		Response.Write( "<td align=left>" & strA & wk_item &"</a></td>")
		Response.Write( "</tr>")
	'����U�@���O��
		rstObj1.MoveNext
		if rstObj1.EOF=True then exit for
	next	
%>
</table>
<%
end if
'������ƶ�
rstObj1.Close
'���]����ܼ� 
set rstObj1=Nothing
%>

<%
'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing 
%>


</form>

</body>
</html>

