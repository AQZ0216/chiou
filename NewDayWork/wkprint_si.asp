<%@ Language=VBScript CODEPAGE=950 %>
<%
	'讀取人員姓名
	worker = Session("worker")
	wk_id=Request("wk_id")
%>
<!-- 開啟資料庫 -->
<!-- #Include file = "./include/opendb_daywork.inc" -->
<%
'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_id =" & wk_id
rstObj1.open strSQL_show,conDB,3,1
'讀取資料
undo_date1=rstObj1.fields("undo_date1")
doing_date1=rstObj1.fields("doing_date1")
done_date1=rstObj1.fields("done_date1")
wk_item=rstObj1.fields("wk_item")
wk_content=rstObj1.fields("wk_content")
wk_order=rstObj1.fields("wk_order")
wk_doer=rstObj1.fields("wk_doer")
wk_checker=rstObj1.fields("wk_checker")
wk_undoer=rstObj1.fields("wk_undoer")
wk_class=rstObj1.fields("wk_class")
%>
<%
'關閉資料集
rstObj1.Close
'重設資料變數 
set rstObj1=Nothing
'關閉資料庫 
conDB.Close
'重設物件變數 
set conDB=Nothing 

'將換行字元chr(13)轉成html中之換行標籤<br>
if isnull(wk_content) then
	wk_contenta="(空白)"
else
	wk_contenta=replace(wk_content,chr(13),"<br>")
end if

%>

<html>
<head>
<title>工作列印</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="icon" href="../daywork/img/khouse.ico" type="image/ico" />
<style type="text/css"><!--
body{font-family:'新細明體';background-color :'#FFFEEE'}

td{
	margin:5px 0 0 0;		/*邊緣上下左右*/
	border-color:'#000000'; /*表格外框顏色*/ 
	border-style:solid;		/*表格外框線型*/
	border-width:1px;		/*表格外框厚度*/  
	vertical-align:middle;	/*字體垂直對齊方式*/
	}
table{	
	/*margin:0 0 0 0;		/*邊緣上下左右*/
	border-collapse:collapse; 	/*邊框形式重合*/
	}
--></style>
</head>
<body>
<center>
<font size=5 color='blue'>工作列印</font><br>
<table border=0 cellspacing=0 cellpadding=2 style="width:600px" >
<col style="width:16%;background-color:#d3d3d3;" align=right>
<col style="width:16%;" align=left>
<col style="width:16%;background-color:#d3d3d3;" align=right>
<col style="width:16%;" align=left>
<col style="width:16%;background-color:#d3d3d3;" align=right>
<col style="width:16%;" align=left>
<tr>
	<td>工作編號：
	<td><%=wk_id%>
	<td>公告日期：
	<td><%=undo_date1%>
	<td>執行日期：
	<td><%=doing_date1%>
<tr>
	<td>主旨：</td>
	<td colspan=5 ><%=wk_item%>
<tr>
	<td colspan=6 style="text-align:center;background-color:#d3d3d3;">執行內容：
<tr>
	<td colspan=6 style="text-align:left;background-color:#FFFEEE;padding: 6px 6px 6px 12px;" >
	
	<%=wk_contenta%><br>
	
</table>
</center>

</body>

</html>
