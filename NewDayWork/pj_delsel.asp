<%@ Language=VBScript CODEPAGE=950 %>
<%
botton_color="#c3c3c3"
%>
<%
wk_id=request("wk_id")
pj_id=request("p_id")
%>
<%
' �s��Access��Ʈw./database/daywork.mdb
DBpath=Server.MapPath("./database/daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="work_data"

'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_id =" & wk_id
rstObj1.open strSQL_show,conDB,3,3
'Ū�����
wk_item=rstObj1.fields("wk_item")

'������ƶ�
rstObj1.Close
'���]����ܼ� 
set rstObj1=Nothing
'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing 
%>	 
<html>
<head>
<title>�M�׸�Ƨ�</title>
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
<form name="form1" method=post action="pj_delsel_ok.asp">
<input type=hidden name="p00" value="<%=wk_id%>">
<font style="font-size:20px;font-weight:bold;letter-spacing:15px;">�����i�M�צW�١j</font><br>
<table border=0 cellspacing=0 cellpadding=2 style="width:660px" >
<col style="width:100px;" align=right>
<col style="" align=left>
<tr style="height:25px;">
<td colspan=2 align='center'>
	<input class=imenu type=submit name=sentb value="�T�w" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='<%=botton_color%>';">
	<input class=imenu type=button name=giveup value="�^�W�@��" onclick="history.back()"  onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='<%=botton_color%>';">
</td></tr>
<tr>
<td><font color="red">�u�@�s���G </font></td>
<td>
<%=wk_id%>
</td>
</tr>
<tr>
<td><font color="red">�u�@�D���G </font></td>
<td>
<%=wk_item%>
</td>
</tr>
<tr>
<td style="vertical-align:top;"><font color="red">�M�צW�١G </font></td>
<td>
<%
' �s��Access��Ʈw./database/daywork.mdb
DBpath=Server.MapPath("./database/daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="project_data"

'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where pj_id="&pj_id
rstObj1.open strSQL_show,conDB,3,3
p_id=rstObj1.fields("pj_id")	'�M��id
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
%>
��<%=p_01%>����<%=p_02%>�� 
</td>
</tr>
</table>
</form>

</body>
</html>

