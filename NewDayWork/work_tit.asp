<% @codepage=950%>
<%
	'讀取人員姓名
	worker = Session("worker")
'	worker = request("worker")
%>
<!-- #Include file = "./include/array_worker_crew.inc" -->
<%
for i=1 to worker_no
	if worker=worker_a(i-1) then
	  pwkr_id=staff_id_a(i-1)
	else 
	end if
next	 
%>
<HTML>
<HEAD>
<title>工作管理系統</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'微軟正黑體';background-color:'#F0FFF0';margin-top:10;}
input{font-family:'微軟正黑體';font-size:12pt;}
textarea{font-family:'微軟正黑體';}
SELECT{font-family:'微軟正黑體';font-size:12pt;}
td{font-family:'微軟正黑體';}
--></style>
</HEAD>
<BODY>
<!-- 標題開始 -->

<CENTER>
	<FORM name="form1" action="" method=post >
<table border=0 cellspacing=0 cellpadding=0 >
<col >
<tr>
    <td colspan=1 align=center>
	<!-- Include file = "./include/toolbar_worker_tit.inc" -->
<input style="cursor:hand;width:90px;background-color:'#d3d3d3';" type="button"  name="bk" value="回首頁" onclick="parent.location.href='firstpage.asp'" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
<font style="font-size:16pt;color:blue;font-weight:bold;letter-spacing:2pt;background-color:#eeeeee;">
&nbsp;【<%=worker%>】工作管理&nbsp;
	</font>&nbsp;&nbsp;
			<SELECT name="chgmen_w" onchange="changeworker()">
		<option value="" selected>人員更換</option>
			<%
				for i=1 to worker_no
					response.write "<option value='" & worker_a(i-1) & "'>" & worker_a(i-1) &"</option>"
				next
			%>
		</SELECT>
&nbsp;&nbsp;
	<img src="./img/clock.png" style="height:30px;vertical-align:middle;cursor:hand;" title="查詢出勤時間" onclick="querytime()" >
	</td>
</tr>
<tr>
<td colspan=1 align=center>
	<!-- #Include file = "./include/toolbar_work_tit.inc" -->
</td>
</tr>
</table>
</form>
<!--<hr width=800> -->
<!-- 標題結束 -->

</center>
<script language=vbscript>	
sub querytime () '查詢刷卡時間
		   MyVar = MsgBox ("確定查詢刷卡時間！！。", 64+1, "MsgBox Example")
		   if MyVar =1 then
		   	'確定編輯
		   	'window.open  ""
		   end if

end sub	
sub changeworker () '登入選擇人員
	ppworker=document.form1.chgmen_w.value
	if ppworker="" then
		MyVar = MsgBox ("請選擇人員！！。", 64+0, "MsgBox Example")
	else
		parent.location.href="./work_main.asp?worker="&ppworker
	end if
end sub	
</script>
</BODY>
</HTML>
