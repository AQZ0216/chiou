<%@ Language=VBScript CODEPAGE=950 %>
<%
wk_id=request("wk_id")
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
<font style="font-size:5mm;color:blue;">附加檔案功能</font>
<table border=1 cellspacing=0 cellpadding=0>
<col width=100>
<col width=400>
<tr>
	<td align="right">
	<font color="red">工作編號：</font>
	</td>
	<td >
	<input type='text' name='wk_id' value='<%=wk_id%>' readonly style="width:20%;">
	</td>
</tr><tr>
	<td align="right">
	<font color="red">檔案名稱：</font>
	</td>
	<td >
	<input type='file' name='filename' value='' style="width:100%;">
	</td>
</tr>
<tr>
	<td colspan=2 align="center">
	<input type="submit" name="press" value="確定新增" >
	<input type="reset" name="cancel" value="清除資料" >
	</td>
<tr>
</table>

<hr color=red>
因網路權限管理問題，檔案並沒有真正存入指定的目錄中。<br>
此功能僅是將檔案名稱附加進資料庫中，以供連結檔案之用。<br>
請先自行將檔案存入指定目錄中。<br>
目錄：<font color=blue>網路上的芳鄰//chiou-server/d/chiou/daywork/addfile </font>。<br> 
<hr color=red>

</form>
</center>
</body>
</html>
