<%@ Language=VBScript CODEPAGE=950 %>
<%
row = request("row")
col = request("col")
%>
<%
'Ū���������O�}�C
' �s��Access��Ʈw./database/linkweb.mdb
DBpath=Server.MapPath("./database/linkweb.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="linkdata"
'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
   strSQL_show="Select * from " & tb_name & " where lk_row="& row &" and lk_col="& col &" order by lk_id asc"
rstObj1.open strSQL_show,conDB,3,1
'�p�����`��	
totalput01=rstObj1.recordcount
'�C�X��ƶ���
      p_id=rstObj1.fields("lk_id")		'id	
      p_01=rstObj1.fields("lk_url")		'�s�����}
      p_02=rstObj1.fields("lk_item")		'�u���D
      p_03=rstObj1.fields("lk_title")		'�y�z
      p_04=rstObj1.fields("lk_row")		'�C
      p_05=rstObj1.fields("lk_col")		'��

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
<title>�ק�s�����</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--�]�w�˪O�榡-->
<style type="text/css">
	<!--
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
	font-size:15px;				/*�r��j�p*/
	font-weight:bold;
	cursor:hand;				/*��ЧΦ�*/
	background-color:'<%=botton_color%>'; 		
	margin:0 0 0 0;		/*��t�W�U���k*/
	width:100px;
	height:100%;
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
	font-size:12pt;
	}
table{	
	margin:0 0 0 0;		/*��t�W�U���k*/
	border-collapse:collapse; 	/*��اΦ����X*/
	}
--></style>
<body bgcolor="#ececec" style="margin:10 0 0 0;padding:10 0 0 0;">
<center>
<form name="form_ipt" method=post action="0_edit_link_ok.asp">
<input type=hidden name="p_id" value="<%=p_id%>" >
<table border=0 cellspacing=0 cellpadding=0 style="width:660px" >
<col style="width:100pt;text-align:center;">
<col style="width:550pt;text-align:left;padding-left:2pt;">
<tr style="height:25pt;">
	<td colspan=6 style="font-size:15pt;font-weight:bold;letter-spacing:5pt;">�s����ƭק�</td>
</tr>
<tr style="height:25px;"><td colspan=2 align='center'>
      <input type=submit name=sent value="�T�w�ק�" >
      <input type=button name=giveup value="�^�W�@��" onclick="history.back()" >
</td></tr>
<tr style="height:25pt;">
	<td style="color:red;">�s�����}</td>
	<td ><input type='text' name="p_01" value="<%=p_01%>" style="width:99%;" ></td>
</tr>
<tr style="height:25pt;">
	<td style="color:red;">²�u���D</td>
	<td colspan=5 ><input type='text' name="p_02" value="<%=p_02%>" style="width:100pt;" maxlength='14'>[����r���q�b7�r�����C���W�L8�r]</td>
</tr>
<tr style="height:25pt;">
	<td style="color:red;">�y�z</td>
	<td colspan=5 ><input type='text' name="p_03" value="<%=p_03%>" style="width:99%;" ></td>
</tr>
</table>
</form>
</center>	
</body>
</html>

