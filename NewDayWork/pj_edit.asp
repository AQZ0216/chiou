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
'if p_id="" or isnull(p_id) then
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
'if pj_id="" or isnull(p_id) then
'	strSQL_show="Select * from " & tb_name & " where pj_02 =" & pj_02
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
<title>�M�׸��</title>
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
<form name="form1" method=post action="pj_edit_ok.asp">
<input type=hidden name="p_id" value="<%=p_00%>">
<font style="font-size:20px;font-weight:bold;letter-spacing:15px;">�i�M�צW�١j���</font><br>
<table border=0 cellspacing=0 cellpadding=2 style="width:660px" >
<col style="width:100px;" align=right>
<col style="" align=left>
<tr style="height:25px;">
<td colspan=2 align='center'>
	<input class=imenu type=submit name=sentb value="�T�w�ק�" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='<%=botton_color%>';">
	<input class=imenu type=reset name=reset value="�M�����"  onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='<%=botton_color%>';">
	<input class=imenu type=button name=giveup value="�^�W�@��" onclick="history.back()"  onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='<%=botton_color%>';">
</td></tr>
<tr>
<td><font color="red">�M�׽s���G </font></td>
<td>
<input class=itext type=text name="p01" value="<%=p_01%>" maxlength="10" style="width:120px;">
���M�׽s������10�Ӧr�զ��C</td>
</tr>
<tr>
<td><font color="red">�M�צW�١G </font></td>
<td>
<input class=itext type=text name="p02" value="<%=p_02%>" maxlength="10" style="width:120px;">
���M�צW�١���10�Ӧr�զ��C</td>
</tr>
<tr>
<td><font color="red">�M�׻����G </font></td>
<td>
<textarea class=itext name="p03" style="height:50px;width:100%;background-color:'#ffeedd';" ><%=p_03%></textarea>
���M�׻������i�ק�C</td>
</tr>
</table>
</form>

</body>
</html>

