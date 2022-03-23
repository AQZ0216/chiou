<%@ Language=VBScript CODEPAGE=950 %>
<%
botton_color="#c3c3c3"
%>
<%
wk_id=request("wk_id")
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
<form name="form1" method=post action="pj_add_ok.asp">
<input type=hidden name="wk_id" value="<%=wk_id%>">
<font style="font-size:20px;font-weight:bold;letter-spacing:15px;">【專案名稱】新增</font><br>
<table border=0 cellspacing=0 cellpadding=2 style="width:660px" >
<col style="width:100px;" align=right>
<col style="" align=left>
<tr style="height:25px;">
<td colspan=2 align='center'>
	<input class=imenu type=submit name=sentb value="確定新增" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='<%=botton_color%>';">
	<input class=imenu type=reset name=reset value="清除資料"  onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='<%=botton_color%>';">
	<input class=imenu type=button name=giveup value="回上一頁" onclick="history.back()"  onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='<%=botton_color%>';">
</td></tr>
<%
if wk_id="" or isnull(wk_id) then
else
%>
<!-- 開啟資料庫 -->
<!-- #Include file = "./include/opendb_daywork.inc" -->
<%
'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_id =" & wk_id
rstObj1.open strSQL_show,conDB,3,1
'讀取資料
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
'關閉資料集
rstObj1.Close
'重設資料變數 
set rstObj1=Nothing
'關閉資料庫 
conDB.Close
'重設物件變數 
set conDB=Nothing 
%>
<tr>
<td><font color="red">工作編號： </font></td>
<td>
<input class=itext type=text name="p00" value="<%=wk_id%>" style="width:120px;" readonly>
</td>
</tr>
<tr>
<td><font color="red">工作主旨： </font></td>
<td>
<%=wk_item%>
</td>
</tr>
<%
end if
%>
<tr>
<td><font color="red">專案編號： </font></td>
<td>
<input class=itext type=text name="p01" value="" maxlength="10" style="width:120px;" >
﹝專案編號﹞為10個字組成。</td>
</tr>
<tr>
<td><font color="red">專案名稱： </font></td>
<td>
<input class=itext type=text name="p02" value="" maxlength="10" style="width:120px;" >
﹝專案名稱﹞為10個字組成。</td>
</tr>
<tr>
<td><font color="red">專案說明： </font></td>
<td>
<textarea class=itext name="p03" style="height:50px;width:100%;background-color:'#ffeedd';" ></textarea>
</td>
</tr>
</table>
</form>

</body>
</html>

