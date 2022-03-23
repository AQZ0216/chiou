<%@ Language=VBScript CODEPAGE=950 %>
<%
botton_color="#c3c3c3"
%>
<%
wk_id=request("wk_id")
pj_id=request("p_id")
%>
<%
' 連結Access資料庫./database/daywork.mdb
DBpath=Server.MapPath("./database/daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="work_data"

'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_id =" & wk_id
rstObj1.open strSQL_show,conDB,3,3
'讀取資料
wk_item=rstObj1.fields("wk_item")

'關閉資料集
rstObj1.Close
'重設資料變數 
set rstObj1=Nothing
'關閉資料庫 
conDB.Close
'重設物件變數 
set conDB=Nothing 
%>	 
<html>
<head>
<title>專案資料夾</title>
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
	font-size:15px;
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
     }

--></style>
</head>
<body>
<center>
<form name="form1" method=post action="pj_delsel_ok.asp">
<input type=hidden name="p00" value="<%=wk_id%>">
<font style="font-size:20px;font-weight:bold;letter-spacing:15px;">取消【專案名稱】</font><br>
<table border=0 cellspacing=0 cellpadding=2 style="width:660px" >
<col style="width:100px;" align=right>
<col style="" align=left>
<tr style="height:25px;">
<td colspan=2 align='center'>
	<input class=imenu type=submit name=sentb value="確定" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='<%=botton_color%>';">
	<input class=imenu type=button name=giveup value="回上一頁" onclick="history.back()"  onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='<%=botton_color%>';">
</td></tr>
<tr>
<td><font color="red">工作編號： </font></td>
<td>
<%=wk_id%>
</td>
</tr>
<tr>
<td><font color="red">工作主旨： </font></td>
<td>
<%=wk_item%>
</td>
</tr>
<tr>
<td style="vertical-align:top;"><font color="red">專案名稱： </font></td>
<td>
<%
' 連結Access資料庫./database/daywork.mdb
DBpath=Server.MapPath("./database/daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="project_data"

'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where pj_id="&pj_id
rstObj1.open strSQL_show,conDB,3,3
p_id=rstObj1.fields("pj_id")	'專案id
p_01=rstObj1.fields("pj_01")	'專案編號
p_02=rstObj1.fields("pj_02")	'專案名稱
p_03=rstObj1.fields("pj_03")	'專案說明

'關閉資料集
rstObj1.Close
'重設資料變數 
set rstObj1=Nothing
'關閉資料庫 
conDB.Close
'重設物件變數 
set conDB=Nothing 
%>
﹝<%=p_01%>﹞﹝<%=p_02%>﹞ 
</td>
</tr>
</table>
</form>

</body>
</html>

