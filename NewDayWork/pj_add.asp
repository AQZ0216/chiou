<%@ Language=VBScript CODEPAGE=950 %>
<%
botton_color="#c3c3c3"
%>
<%
wk_id=request("wk_id")
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
<form name="form1" method=post action="pj_add_ok.asp">
<input type=hidden name="wk_id" value="<%=wk_id%>">
<font style="font-size:20px;font-weight:bold;letter-spacing:15px;">�i�M�צW�١j�s�W</font><br>
<table border=0 cellspacing=0 cellpadding=2 style="width:660px" >
<col style="width:100px;" align=right>
<col style="" align=left>
<tr style="height:25px;">
<td colspan=2 align='center'>
	<input class=imenu type=submit name=sentb value="�T�w�s�W" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='<%=botton_color%>';">
	<input class=imenu type=reset name=reset value="�M�����"  onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='<%=botton_color%>';">
	<input class=imenu type=button name=giveup value="�^�W�@��" onclick="history.back()"  onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='<%=botton_color%>';">
</td></tr>
<%
if wk_id="" or isnull(wk_id) then
else
%>
<!-- �}�Ҹ�Ʈw -->
<!-- #Include file = "./include/opendb_daywork.inc" -->
<%
'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_id =" & wk_id
rstObj1.open strSQL_show,conDB,3,1
'Ū�����
'undo_date1=rstObj1.fields("undo_date1")
'doing_date1=rstObj1.fields("doing_date1")
'done_date1=rstObj1.fields("done_date1")
wk_item=rstObj1.fields("wk_item")
'wk_content=rstObj1.fields("wk_content")
'wk_order=rstObj1.fields("wk_order")
'wk_doer=rstObj1.fields("wk_doer")
'wk_checker=rstObj1.fields("wk_checker")
'wk_undoer=rstObj1.fields("wk_undoer")
'wk_class=rstObj1.fields("wk_class")
'wk_group=rstObj1.fields("wk_group")
'wk_exe=rstObj1.fields("wk_exe")
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
%>
<tr>
<td><font color="red">�u�@�s���G </font></td>
<td>
<input class=itext type=text name="p00" value="<%=wk_id%>" style="width:120px;" readonly>
</td>
</tr>
<tr>
<td><font color="red">�u�@�D���G </font></td>
<td>
<%=wk_item%>
</td>
</tr>
<%
end if
%>
<tr>
<td><font color="red">�M�׽s���G </font></td>
<td>
<input class=itext type=text name="p01" value="" maxlength="10" style="width:120px;" >
���M�׽s������10�Ӧr�զ��C</td>
</tr>
<tr>
<td><font color="red">�M�צW�١G </font></td>
<td>
<input class=itext type=text name="p02" value="" maxlength="10" style="width:120px;" >
���M�צW�١���10�Ӧr�զ��C</td>
</tr>
<tr>
<td><font color="red">�M�׻����G </font></td>
<td>
<textarea class=itext name="p03" style="height:50px;width:100%;background-color:'#ffeedd';" ></textarea>
</td>
</tr>
</table>
</form>

</body>
</html>

