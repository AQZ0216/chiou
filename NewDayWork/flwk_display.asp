<%@ Language=VBScript CODEPAGE=950 %>

<%
fw_id=request("fw_id")
' �s��Access��Ʈw./database/daywork.mdb
DBpath=Server.MapPath("./database/daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="wk_file"

'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where fw_id =" & fw_id
rstObj1.open strSQL_show,conDB,3,1

wk_id=rstObj1.fields("wk_id")		'�����_
fl_name=trim(rstObj1.fields("fl_name"))	'�Ȥ�id

'������ƶ�
rstObj1.Close
'���]����ܼ� 
set rstObj1=Nothing

'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing 

for i=1 to 10
  pos1=instr(1,fl_name,"\",1)
  fl_name=right(fl_name,len(fl_name)-pos1)
  if pos1=0 then exit for
next


%>
<HTML>
<HEAD>
<title>�u�@�޲z�t��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
body {  scrollbar-3dlight-color:#ffffff;
        scrollbar-arrow-color:#CCCCCC;
        scrollbar-base-color:#666633;
        scrollbar-darkshadow-color:#e6e6cc;
        scrollbar-face-color:#666666;
        scrollbar-highlight-color:#ffffff;
        scrollbar-shadow-color:#e6e6cc;
        scrollbar-track-color:#ffffff;
        margin:2mm 0mm 0mm 0mm;		/*��t�W�U���k*/
		font-family:'�з���';		/*�r��*/
		font-size:4mm; 			/*�r��j�p*/
		background-color:'#F0FFF0';
     }
input.imenu { 
	font-size:3.5mm;				/*�r��j�p*/
	cursor:hand;				/*��ЧΦ�*/ 
	background-color:'#d3d3d3'; 		/*�~���C��*/ 
	margin:0 0 0 0;		/*��t�W�U���k*/
     }
input.imenu1 { 
	font-size:3.5mm;	/*�r��j�p*/
	font-weight:bold;				
	cursor:hand;				/*��ЧΦ�*/ 
	background-color:'#eeeeff'; 		/*�~���C��*/ 
	margin:0 0 0 0;		/*��t�W�U���k*/
	width:80px;
	height:100%;
     }
     
TD.SOME{
		font-family: '�з���';
		font-size: 3.3mm;
		line-height: 18px;
		color:blue;
		font-weight:bold;
		}
TD.myd{
		font-family: '�з���';
		font-size: 3.3mm;
		line-height: 18px;
		background-color:#f0ffff;
		}     
    
-->
</style>

</HEAD>
<BODY>
<center>

<form name="form1" action="flwk_add_ok.asp" method="post" >
<table border=1 cellspacing=0 cellpadding=0>
<col width=100>
<col width=400>
<tr>
	<td align="right">
	<font color="red">�u�@�s���G</font>
	</td>
	<td >
	<%=wk_id%>
	</td>
</tr><tr>
	<td align="right">
	<font color="red">�ɮצW�١G</font>
	</td>
	<td >
	<a href="http://192.168.123.112/addfile/<%=fl_name%>" target="_blank" ><%=fl_name%></a>
	</td>
</tr>
</table>
</form>

</center>
</body>
</html>
