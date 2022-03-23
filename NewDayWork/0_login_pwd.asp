<%@ Language=VBScript CODEPAGE=950 %>
<%
   '讀取資料帳號、密碼
   p_worker=request("worker")
   wkr_pwd=request("wkr_pwd")
if p_worker="" or isnull(p_worker) then
      str_url="./firstpage.asp"
      response.redirect(str_url)      '轉址到首頁
else
   if wkr_pwd="" or isnull(wkr_pwd) then
   else
       session("wkr_pwd")=wkr_pwd
      str_url="./work_main.asp?worker="&p_worker
      response.redirect(str_url)      '轉址到首頁
   end if

end if
%>

<html>
<head>
<title>密碼檢查</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'新細明體';background-color :'#FFFEEE'}
input{
	font-family:'新細明體';
	font-size:12pt;
	}
select{font-family:'新細明體';font-size:10pt;cursor:hand;}
.itxt{
	font-family:'新細明體';
	font-size:12pt;
	width:100%;
	height:100%;
	}
input.imenu { 
	/*font-size:15px;				/*字體大小*/
	/*font-weight:bold;
	cursor:hand;				/*游標形式*/ 
	background-color:'<%=botton_color%>'; 		
	margin:0 0 0 0;		/*邊緣上下左右*/
	width:100px;
	/*height:100%;*/
	color:#000000;
	letter-spacing:2px;
	cursor:hand;
     }
td{
	margin:0 0 0 0;		/*邊緣上下左右*/
	border-color:'#808080'; /*表格外框顏色*/ 
	border-style:solid;		/*表格外框線型*/
	border-width:1px;		/*表格外框厚度*/  
	vertical-align:middle;	/*字體垂直對齊方式*/
	/*font-size:15px;*/ 
	}
table{	
	margin:0 0 0 0;		/*邊緣上下左右*/
	border-collapse:collapse; 	/*邊框形式重合*/
	}
input.itext { 
	font-size:3.5mm;				/*字體大小*/
	/*cursor:hand;				/*游標形式*/ 
	width:100%;
	height:5mm;
	background-color:'#ffeedd'; 		/*外框顏色*/
	margin:0 0 0 0;		/*邊緣上下左右*/
	color:black;
	text-align:right;
     }

--></style>
</head>
<body>
<center>
<form name="form_login" method=post action="0_login_pwd.asp">
<input type=hidden name="worker" value="<%=p_worker%>" >
<font style="font-size:20px;font-weight:bold;letter-spacing:15px;">【登入工作管理系統】</font>
<!-- [<%=p_userid%>][<%=p_pwd%>] -->
<br>
<hr color=red> 
<table border=0 cellspacing=0 cellpadding=2 style="width:300px" >
<col style="width:100px;font-size:4mm;" align=center>
<col style="width:200px;font-size:4mm;" align=center>
<tr>
<td>使用者：</td>
<td><%=p_worker%></td>
</tr>
<tr>
<td>密碼：</td>
<td><input type="password" style="text-align:left;" name="wkr_pwd" value="" ></td>
</tr>
<tr>
<td colspan=2>
	<input type="submit" name="submit01" value="確定" >
	<input type="reset" name="reset01" value="重設" >
</td>
</tr>
</table> 
</body>
</html>

