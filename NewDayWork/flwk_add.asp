<%@ Language=VBScript CODEPAGE=950 %>
<%
wk_id=request("wk_id")
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
<font style="font-size:5mm;color:blue;">���[�ɮץ\��</font>
<table border=1 cellspacing=0 cellpadding=0>
<col width=100>
<col width=400>
<tr>
	<td align="right">
	<font color="red">�u�@�s���G</font>
	</td>
	<td >
	<input type='text' name='wk_id' value='<%=wk_id%>' readonly style="width:20%;">
	</td>
</tr><tr>
	<td align="right">
	<font color="red">�ɮצW�١G</font>
	</td>
	<td >
	<input type='file' name='filename' value='' style="width:100%;">
	</td>
</tr>
<tr>
	<td colspan=2 align="center">
	<input type="submit" name="press" value="�T�w�s�W" >
	<input type="reset" name="cancel" value="�M�����" >
	</td>
<tr>
</table>

<hr color=red>
�]�����v���޲z���D�A�ɮרèS���u���s�J���w���ؿ����C<br>
���\��ȬO�N�ɮצW�٪��[�i��Ʈw���A�H�ѳs���ɮפ��ΡC<br>
�Х��ۦ�N�ɮצs�J���w�ؿ����C<br>
�ؿ��G<font color=blue>�����W���ھF//chiou-server/d/chiou/daywork/addfile </font>�C<br> 
<hr color=red>

</form>
</center>
</body>
</html>
