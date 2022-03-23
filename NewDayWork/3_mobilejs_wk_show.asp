<%@ Language=VBScript CODEPAGE=950 %>
<%
	'讀取人員姓名
	worker = Session("worker")
	wk_id=Request("wk_id")
%>
<!-- 開啟資料庫 -->
<!-- Include file = "./include/opendb_daywork.inc" -->
<%
' 連結Access資料庫daywork.mdb
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
wk_group=rstObj1.fields("wk_group")
wk_exe=rstObj1.fields("wk_exe")
wk_pjn=rstObj1.fields("pj_02")   '專案名稱
wk_att=rstObj1.fields("wk_att")           '出席人員

'關閉資料集
rstObj1.Close
'重設資料變數 
set rstObj1=Nothing

'關閉資料庫 
conDB.Close
'重設物件變數 
set conDB=Nothing 
%>
<%
' 連結Access資料庫temp-daywork.mdb
DBpath=Server.MapPath("./database/temp-daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="work_data"
'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where tmp_id ="&wk_id&" and ipt_ok=0 order by wk_id desc" 
rstObj1.open strSQL_show,conDB,1,1
tpn=rstObj1.recordcount
'關閉資料集
rstObj1.Close
'重設資料變數 
set rstObj1=Nothing

'關閉資料庫
conDB.Close
'重設物件變數
set conDB=Nothing
%>
<HTML>
<HEAD>
<title>工作管理系統</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'標楷體';background-color:'#F0FFF0'}
--></style>
</HEAD>
<BODY>
<center>
<form name="form1" action="" method="post">
<input type="hidden" name="wk_id1" value="<%=wk_id%>">
<input type="hidden" name="worker1" value="<%=worker%>">
<input type="hidden" name="wk_order1" value="<%=wk_order%>">
<center>

<table border=0 cellspacing=0 cellpadding=0>
<col span=7 style="width:100px;text-align:center;">
<tr width=720>	
<% if tpn=1 then %>
<td>【未同步】</td>
<% else %>
<td>【已同步】</td>
<% end if %>
	<td>
		<input type=button name="bkpg" value="回上一頁" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="parent.location.href='javascript:history.back()'" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
<% if tpn=1 then %>
	<td> 	<input type=button name="edit" value="編修未同步工作" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="wk_edit()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';"> 	</td>
	<td>	<input type=button name="delete" value="刪除未同步工作" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="wk_del()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">	</td>
<% else %>
	<td>	<input type=button name="delete" value="刪除已同步工作" title="刪除已更新的工作" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="wk_delnext()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">	</td>
<% end if %>
	<td>	<input type=button name="wkprint" value="回日曆表" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="wk_calendar('<%=year(doing_date1)%>','<%=month(doing_date1)%>')"  onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';"></td>

</tr>
</table>
<table border=1 cellspacing=0 cellpadding=0>
<col width=120><col width=120><col width=120><col width=120><col width=120><col width=120>
<tr>
	<td align="center" colspan=2><font size=4 color="red"><b>顯示單一工作表(<%=wk_group%>)</b></font></td>
	<td align="right">工作編號：</td>
	<td><%=wk_id%></td>
	<td align="right">工作分類：</td>
	<td><%=wk_class%></td>

</tr>
<tr>
	<td align="right">公告者：</td>
	<td><%=wk_order%></td>
	<td align="right">公告日期：</td>
	<td><%=undo_date1%></td>
	<td align="right">執行日期：</td>
	<td><%=doing_date1%></td>
</tr>

<tr>
	<td align="right">
	知會人員：
	</td>
	<td colspan=5>
	<%=wk_doer%>
	</td>
</tr>
<tr>
	<td align="right">
	執行人員：
	</td>
	<td colspan=5>
	<%=wk_exe%>
	</td>
</tr>
<tr>
	<td align="right">
	出席人員：
	</td>
	<td colspan=5>
	<%=wk_att%>
	</td>
</tr>
<!--
<tr>
	<td align="right">
	完成人員：
	</td>
	<td colspan=5>
	<%=wk_checker%>
	</td>
</tr>
-->
<!--<tr>
	<td align="right">
	未完成人員：
	</td>
	<td colspan=5>
	<%=wk_undoer%>
	</td>
</tr>-->
<tr>
	<td align="right">
	主旨：
	</td>
	<td colspan=5>
	<%=wk_item%>
	</td>
</tr>
<tr>
	<td align="right" valign="top">
	執行內容：
	</td>
	<td colspan=5>
	<%
	wk_content_s=replace(wk_content,chr(13),"<br>")
	%>
	<font style="font-size:12pt;" ><%=wk_content_s%></font>
	</td>
</tr>

</table>
</center>

<!-- Include file = "./include/wk_show_form.inc" -->

<script type="text/javascript">
function wk_edit()
{
   var x1=document.forms["form1"]["wk_id1"].value;
   str_url="./3_mobilejs_wk_edit.asp?wk_id="+x1;
   // alert(str_url);
   location.href = str_url ;
}
function wk_del()
{
   var x1=document.forms["form1"]["wk_id1"].value;
   str_url="./3_mobilejs_wk_del_ok.asp?wk_id="+x1;
		var r = confirm("請確認是否刪除【未同步】的工作？？"+String.fromCharCode(13,10)+str_url);
		if (r == true) {
		    //txt = "確認刪除！！";
		    //alert(txt);
		    location.href = str_url ;
		} else {
		    //txt = "取消刪除！！";
		    //alert(txt);
		}
   // alert(str_url);
   //location.href = str_url ;
}
function wk_calendar(pyear,pmonth)
{
   var x1=document.forms["form1"]["wk_id1"].value;
   str_url="./wk_calendar_all.asp?nYear="+pyear+"&nMonth="+pmonth;
   // alert(str_url);
   location.href = str_url ;
}
function wk_delnext()
{
   var x1=document.forms["form1"]["wk_id1"].value;
   str_url="./3_mobilejs_wk_delnext_ok.asp?wk_id="+x1;
		var r = confirm("請確認是否刪除【已同步】的工作？？"+String.fromCharCode(13)+str_url);
		if (r == true) {
		    //txt = "確認刪除！！";
		    //alert(txt);
		    location.href = str_url ;
		} else {
		    //txt = "取消刪除！！";
		    //alert(txt);
		}
   // alert(str_url);
   //location.href = str_url ;
}
</script>
</form>
</center>

</body>
</html>
