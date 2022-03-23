<%@ Language=VBScript CODEPAGE=950 %>

<%
fw_id=request("fw_id")
' 連結Access資料庫./database/daywork.mdb
DBpath=Server.MapPath("./database/daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="wk_file"

'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where fw_id =" & fw_id
rstObj1.open strSQL_show,conDB,3,1

wk_id=rstObj1.fields("wk_id")		'完成否
fl_name=trim(rstObj1.fields("fl_name"))	'客戶id

'關閉資料集
rstObj1.Close
'重設資料變數 
set rstObj1=Nothing

'關閉資料庫 
conDB.Close
'重設物件變數 
set conDB=Nothing 

for i=1 to 10
  pos1=instr(1,fl_name,"\",1)
  fl_name=right(fl_name,len(fl_name)-pos1)
  if pos1=0 then exit for
next


%>
<HTML>
<HEAD>
<title>工作管理系統</title>
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
        margin:2mm 0mm 0mm 0mm;		/*邊緣上下左右*/
		font-family:'標楷體';		/*字形*/
		font-size:4mm; 			/*字體大小*/
		background-color:'#F0FFF0';
     }
input.imenu { 
	font-size:3.5mm;				/*字體大小*/
	cursor:hand;				/*游標形式*/ 
	background-color:'#d3d3d3'; 		/*外框顏色*/ 
	margin:0 0 0 0;		/*邊緣上下左右*/
     }
input.imenu1 { 
	font-size:3.5mm;	/*字體大小*/
	font-weight:bold;				
	cursor:hand;				/*游標形式*/ 
	background-color:'#eeeeff'; 		/*外框顏色*/ 
	margin:0 0 0 0;		/*邊緣上下左右*/
	width:80px;
	height:100%;
     }
     
TD.SOME{
		font-family: '標楷體';
		font-size: 3.3mm;
		line-height: 18px;
		color:blue;
		font-weight:bold;
		}
TD.myd{
		font-family: '標楷體';
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
	<font color="red">工作編號：</font>
	</td>
	<td >
	<%=wk_id%>
	</td>
</tr><tr>
	<td align="right">
	<font color="red">檔案名稱：</font>
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
