<%@ Language=VBScript CODEPAGE=950 %>
<%
   'Ū����Ʊb���B�K�X
   p_worker=request("worker")
   wkr_pwd=request("wkr_pwd")
if p_worker="" or isnull(p_worker) then
      str_url="./firstpage.asp"
      response.redirect(str_url)      '��}�쭺��
else
   if wkr_pwd="" or isnull(wkr_pwd) then
   else
       session("wkr_pwd")=wkr_pwd
      str_url="./work_main.asp?worker="&p_worker
      response.redirect(str_url)      '��}�쭺��
   end if

end if
%>

<html>
<head>
<title>�K�X�ˬd</title>
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
	/*font-size:15px;*/ 
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
	text-align:right;
     }

--></style>
</head>
<body>
<center>
<form name="form_login" method=post action="0_login_pwd.asp">
<input type=hidden name="worker" value="<%=p_worker%>" >
<font style="font-size:20px;font-weight:bold;letter-spacing:15px;">�i�n�J�u�@�޲z�t�Ρj</font>
<!-- [<%=p_userid%>][<%=p_pwd%>] -->
<br>
<hr color=red> 
<table border=0 cellspacing=0 cellpadding=2 style="width:300px" >
<col style="width:100px;font-size:4mm;" align=center>
<col style="width:200px;font-size:4mm;" align=center>
<tr>
<td>�ϥΪ̡G</td>
<td><%=p_worker%></td>
</tr>
<tr>
<td>�K�X�G</td>
<td><input type="password" style="text-align:left;" name="wkr_pwd" value="" ></td>
</tr>
<tr>
<td colspan=2>
	<input type="submit" name="submit01" value="�T�w" >
	<input type="reset" name="reset01" value="���]" >
</td>
</tr>
</table> 
</body>
</html>

