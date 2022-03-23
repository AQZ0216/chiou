<%@ Language=VBScript CODEPAGE=950 %>
<%
row = request("row")
col = request("col")
%>
<%
'讀取分區類別陣列
' 連結Access資料庫./database/linkweb.mdb
DBpath=Server.MapPath("./database/linkweb.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="linkdata"
'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
   strSQL_show="Select * from " & tb_name & " where lk_row="& row &" and lk_col="& col &" order by lk_id asc"
rstObj1.open strSQL_show,conDB,3,1
'計算資料總數	
totalput01=rstObj1.recordcount
'列出資料項目
      p_id=rstObj1.fields("lk_id")		'id	
      p_01=rstObj1.fields("lk_url")		'連結網址
      p_02=rstObj1.fields("lk_item")		'短標題
      p_03=rstObj1.fields("lk_title")		'描述
      p_04=rstObj1.fields("lk_row")		'列
      p_05=rstObj1.fields("lk_col")		'欄

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
<title>修改連結資料</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--設定樣板格式-->
<style type="text/css">
	<!--
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
	font-size:15px;				/*字體大小*/
	font-weight:bold;
	cursor:hand;				/*游標形式*/
	background-color:'<%=botton_color%>'; 		
	margin:0 0 0 0;		/*邊緣上下左右*/
	width:100px;
	height:100%;
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
	font-size:12pt;
	}
table{	
	margin:0 0 0 0;		/*邊緣上下左右*/
	border-collapse:collapse; 	/*邊框形式重合*/
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
	<td colspan=6 style="font-size:15pt;font-weight:bold;letter-spacing:5pt;">連結資料修改</td>
</tr>
<tr style="height:25px;"><td colspan=2 align='center'>
      <input type=submit name=sent value="確定修改" >
      <input type=button name=giveup value="回上一頁" onclick="history.back()" >
</td></tr>
<tr style="height:25pt;">
	<td style="color:red;">連結網址</td>
	<td ><input type='text' name="p_01" value="<%=p_01%>" style="width:99%;" ></td>
</tr>
<tr style="height:25pt;">
	<td style="color:red;">簡短標題</td>
	<td colspan=5 ><input type='text' name="p_02" value="<%=p_02%>" style="width:100pt;" maxlength='14'>[中文字儘量在7字之內。不超過8字]</td>
</tr>
<tr style="height:25pt;">
	<td style="color:red;">描述</td>
	<td colspan=5 ><input type='text' name="p_03" value="<%=p_03%>" style="width:99%;" ></td>
</tr>
</table>
</form>
</center>	
</body>
</html>

