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
'設定讀取資料編號
'if pj_id="" or isnull(pj_id) then
'	Session("flstrbackURL")="pj_show.asp?pj_02="&pj_02
'else
	Session("flstrbackURL")="pj_show.asp?p_id="&pj_id
'end if

'連結Access資料庫./database/daywork.mdb
DBpath=Server.MapPath("./database/daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="project_data"

%>
<%
'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
'if pj_id="" or isnull(pj_id) then
'	strSQL_show="Select * from " & tb_name & " where pj_02 like '"& pj_02 &"'"
'else
	strSQL_show="Select * from " & tb_name & " where pj_id =" & pj_id
'end if
rstObj1.open strSQL_show,conDB,3,1

p_00=rstObj1.fields("pj_id")	'專案id
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

botton_color="#c3c3c3"
%>	
<html>
<head>
<title>刪除專案資料</title>
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
<form name="form1" method=post action="pj_del_ok.asp">
<input type=hidden name="p00" value="<%=p_00%>">
<!--<input type=hidden name="p02" value="<%=p_02%>">-->
<font style="font-size:20px;font-weight:bold;letter-spacing:15px;">【專案名稱】刪除</font><br>
<table border=0 cellspacing=0 cellpadding=2 style="width:760px" >
<col style="width:100px;" align=right>
<col style="" align=left>
<tr style="height:25px;">
<td colspan=2 align='center'>
<font style="color:red;font-size:5mm;font-weight:bold;">
是否確定要刪除專案名稱？&nbsp;&nbsp; 
</font>
	<input class=imenu type=submit name="sent" value="確定刪除"  onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='<%=botton_color%>';" >&nbsp;
	<input class=imenu type=button name=giveup value="回上一頁" onclick="history.back()"  onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='<%=botton_color%>';">
</td></tr>
<tr>
<td><font color="red">專案編號： </font></td>
<td>
<%=p_01%>
</td>
</tr>
<tr>
<td><font color="red">專案名稱： </font></td>
<td>
<%=p_02%>
</td>
</tr>
<tr>
<td><font color="red">專案說明： </font></td>
<td>
<%=p_03%>
</td>
</tr>
</table>

<!--顯示專案與執行工作-->
<%
'連結Access資料庫./database/daywork.mdb
DBpath=Server.MapPath("./database/daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="work_data"
%>
<%
wkgroup="專案工作"
'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_group like '%"&wkgroup&"%' and pj_id ="& pj_id &" order by doing_date1 desc"
rstObj1.open strSQL_show,conDB,3
totalput=rstObj1.recordcount
if totalput=0 then
%>
	<font size=4>無專案工作事項</font>
<%
else
%>
<table border=0 cellspacing=0 cellpadding=2 style="width:760px" >
<col style="width:50px;" align=center>
<col style="width:100px;" align=center>
<col style="" align=left>
<tr style="height:25px;">
	<td colspan=3 align=center>
	<font size=4>所有專案工作事項共:<font color=red><%=totalput%></font>筆</font>
	</td>
</tr>
<tr >
	<td align=center>序號</td>
	<td align=center>執行日期</td>
	<td align=center>主旨</td>
	</td>
</tr>
<%
	'列出資料項目
	rstobj1.MoveFirst
	for i=1 to totalput
	'讀取資料
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
	'移到下一筆記錄
		rstObj1.MoveNext
		if rstObj1.EOF=True then exit for
	next	
%>
</table>
<%
end if
'關閉資料集
rstObj1.Close
'重設資料變數 
set rstObj1=Nothing
%>

<%
wkgroup="一般工作"
'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_group like '%"&wkgroup&"%' and pj_id ="& pj_id &" order by doing_date1 desc"
rstObj1.open strSQL_show,conDB,3
totalput=rstObj1.recordcount
if totalput=0 then
%>
	<font size=4>無一般工作事項</font>
<%
else
%>
<table border=0 cellspacing=0 cellpadding=2 style="width:760px" >
<col style="width:50px;" align=center>
<col style="width:100px;" align=center>
<col style="" align=left>
<tr style="height:25px;">
	<td colspan=3 align=center>
	<font size=4>所有一般工作事項共:<font color=red><%=totalput%></font>筆</font>
	</td>
</tr>
<tr >
	<td align=center>序號</td>
	<td align=center>執行日期</td>
	<td align=center>主旨</td>
	</td>
</tr>
<%
	'列出資料項目
	rstobj1.MoveFirst
	for i=1 to totalput
	'讀取資料
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
	'移到下一筆記錄
		rstObj1.MoveNext
		if rstObj1.EOF=True then exit for
	next	
%>
</table>
<%
end if
'關閉資料集
rstObj1.Close
'重設資料變數 
set rstObj1=Nothing
%>

<%
'關閉資料庫 
conDB.Close
'重設物件變數 
set conDB=Nothing 
%>


</form>

</body>
</html>

